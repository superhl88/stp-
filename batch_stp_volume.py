# -*- coding: utf-8 -*-
import argparse
import json
import multiprocessing
import os
import sys
import time
import traceback
from concurrent.futures import ProcessPoolExecutor, TimeoutError, as_completed
from datetime import datetime
from typing import Any, Dict, List, Optional, Set, Tuple


def check_and_import() -> None:
    missing = []
    try:
        import OCP  # noqa: F401
    except ImportError:
        missing.append("cadquery-ocp")
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        missing.append("openpyxl")
    if missing and not getattr(sys, "frozen", False):
        print("缺少必要的依赖库，正在尝试自动安装:", ", ".join(missing), flush=True)
        import subprocess

        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])


def calc_step_volume(filepath: str) -> Tuple[str, str, str, float, float, int, str]:
    """
    使用 OpenCascade (OCP) 工业级几何内核读取 STEP 文件并精确计算实体体积。
    返回: (文件名(不含后缀), 完整文件名, 所在路径, 体积_mm3, 体积_cm3, 实体数量, 状态)
    """
    filename = os.path.basename(filepath)
    part_no = os.path.splitext(filename)[0]
    try:
        from OCP.BRepGProp import BRepGProp
        from OCP.GProp import GProp_GProps
        from OCP.IFSelect import IFSelect_RetDone
        from OCP.STEPControl import STEPControl_Reader
        from OCP.TopAbs import TopAbs_SOLID
        from OCP.TopExp import TopExp_Explorer

        reader = STEPControl_Reader()
        status = reader.ReadFile(filepath)
        if status != IFSelect_RetDone:
            return (part_no, filename, filepath, 0.0, 0.0, 0, "读取失败(STEP格式异常)")

        reader.TransferRoots()
        shape = reader.OneShape()
        if shape.IsNull():
            return (part_no, filename, filepath, 0.0, 0.0, 0, "空模型(无几何体)")

        # 统计实体(SOLID)数量
        explorer = TopExp_Explorer(shape, TopAbs_SOLID)
        solid_count = 0
        while explorer.More():
            solid_count += 1
            explorer.Next()

        # 计算精确体积
        props = GProp_GProps()
        BRepGProp.VolumeProperties_s(shape, props)
        vol_mm3 = abs(props.Mass())
        vol_cm3 = vol_mm3 / 1000.0

        if vol_mm3 <= 1e-9:
            return (part_no, filename, filepath, 0.0, 0.0, solid_count, "警告: 体积为0(可能仅含曲面/线框)")

        return (part_no, filename, filepath, round(vol_mm3, 4), round(vol_cm3, 6), solid_count, "成功")

    except Exception as e:
        return (part_no, filename, filepath, 0.0, 0.0, 0, f"异常: {str(e)[:80]}")


def _worker_wrapper(filepath: str) -> Tuple[str, str, str, float, float, int, str]:
    """子进程执行包装，确保单文件级安全返回"""
    return calc_step_volume(filepath)


def scan_stp_files(root_dir: str) -> List[str]:
    """递归遍历目标文件夹及其所有子文件夹下的所有 .stp / .step 文件"""
    stp_files: List[str] = []
    skip_dirs = {
        "$RECYCLE.BIN",
        "System Volume Information",
        "Windows",
        "Program Files",
        "Program Files (x86)",
        "ProgramData",
        ".git",
        "__pycache__",
        "build",
        "dist",
    }

    for dirpath, dirnames, filenames in os.walk(root_dir):
        dirnames[:] = [d for d in dirnames if d not in skip_dirs and not d.startswith("$")]
        for fname in filenames:
            if fname.lower().endswith((".stp", ".step")):
                stp_files.append(os.path.join(dirpath, fname))
    stp_files.sort()
    return stp_files


def load_checkpoint(checkpoint_path: str) -> Dict[str, Tuple[str, str, str, float, float, int, str]]:
    """读取历史断点记录，返回 {filepath: record_tuple}"""
    records_by_path: Dict[str, Tuple[str, str, str, float, float, int, str]] = {}
    if not os.path.exists(checkpoint_path):
        return records_by_path

    with open(checkpoint_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
                fp = obj.get("filepath")
                rec = obj.get("record")
                if fp and isinstance(rec, list) and len(rec) == 7:
                    records_by_path[fp] = (
                        str(rec[0]),
                        str(rec[1]),
                        str(rec[2]),
                        float(rec[3]),
                        float(rec[4]),
                        int(rec[5]),
                        str(rec[6]),
                    )
            except Exception:
                # 忽略可能因意外断电截断的最后一行
                continue
    return records_by_path


def append_checkpoint(
    checkpoint_path: str,
    filepath: str,
    record: Tuple[str, str, str, float, float, int, str],
) -> None:
    """实时追加写入单条计算结果到断点文件"""
    payload = {"filepath": filepath, "record": list(record)}
    with open(checkpoint_path, "a", encoding="utf-8") as f:
        f.write(json.dumps(payload, ensure_ascii=False) + "\n")
        f.flush()


def create_pool(workers: int, max_tasks_per_child: int) -> ProcessPoolExecutor:
    """创建进程池，Python 3.11+ 支持 max_tasks_per_child 定期回收内存"""
    kwargs: Dict[str, Any] = {"max_workers": workers}
    if sys.version_info >= (3, 11) and max_tasks_per_child > 0:
        kwargs["max_tasks_per_child"] = max_tasks_per_child
    return ProcessPoolExecutor(**kwargs)


def save_to_excel(records: List[Tuple[str, str, str, float, float, int, str]], output_path: str) -> None:
    """将单卷记录写入格式化良好的 Excel 文件"""
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "STP体积明细"

    headers = ["序号", "料号(文件名)", "完整文件名", "体积 (mm³)", "体积 (cm³)", "实体数量", "状态", "完整路径"]
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws.row_dimensions[1].height = 26

    data_font = Font(name="微软雅黑", size=10)
    warn_font = Font(name="微软雅黑", size=10, color="C00000")

    for idx, rec in enumerate(records, start=1):
        row_data = [idx, rec[0], rec[1], rec[3], rec[4], rec[5], rec[6], rec[2]]
        ws.append(row_data)
        row_num = idx + 1
        is_ok = rec[6] == "成功"
        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=row_num, column=col_idx)
            c.font = data_font if is_ok else warn_font
            c.border = thin_border
            if col_idx in (1, 6, 7):
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in (4, 5):
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.number_format = "#,##0.0000" if col_idx == 4 else "#,##0.000000"
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    col_widths = [8, 28, 32, 18, 18, 12, 24, 60]
    for i, w in enumerate(col_widths, start=1):
        col_letter = openpyxl.utils.get_column_letter(i)
        ws.column_dimensions[col_letter].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{len(records) + 1}"
    wb.save(output_path)


def parse_args(default_dir: str) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="批量计算 STEP 体积（抗崩溃 + 断点续算 + 分卷 Excel）")
    parser.add_argument("--root", type=str, default=default_dir, help="扫描根目录，默认当前运行目录")
    parser.add_argument("--workers", type=int, default=10, help="并行进程数，默认 10（网络盘/大文件建议 8~12）")
    parser.add_argument("--timeout", type=float, default=45.0, help="单文件超时秒数，默认 45s")
    parser.add_argument("--batch-size", type=int, default=200, help="每批提交文件数，默认 200")
    parser.add_argument("--max-tasks-per-child", type=int, default=150, help="子进程处理多少文件后重启（Python 3.11+）")
    parser.add_argument("--checkpoint", type=str, default="", help="断点续算文件路径（默认保存在结果目录内）")
    parser.add_argument("--output-dir", type=str, default="", help="指定输出目录（传入已有目录可继续续算）")
    # 兼容直接双击运行或额外参数，忽略未知参数避免报错
    args, _ = parser.parse_known_args()
    return args


def main() -> None:
    check_and_import()

    if getattr(sys, "frozen", False):
        current_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        current_dir = os.path.dirname(os.path.abspath(__file__))

    args = parse_args(current_dir)
    scan_root = args.root
    workers = max(1, args.workers)
    timeout_sec = max(5.0, args.timeout)
    batch_size = max(10, args.batch_size)
    max_tasks_per_child = max(0, args.max_tasks_per_child)

    print("=" * 60, flush=True)
    print(f"当前运行路径: {current_dir}", flush=True)
    print(f"目标扫描范围: {scan_root}（包含该文件夹及所有子文件夹）", flush=True)
    print(f"运行参数: 进程数={workers} | 单文件超时={timeout_sec}s | 批次={batch_size}", flush=True)
    print("=" * 60, flush=True)
    print("正在递归检索所有 .stp / .step 文件，请稍候...", flush=True)

    stp_files = scan_stp_files(scan_root)
    total_files = len(stp_files)
    print(f"检索完成！共发现 {total_files} 个 STEP 文件。", flush=True)

    if total_files == 0:
        print("在目标目录及子文件夹中未找到任何 .stp 或 .step 文件。")
        return

    if args.output_dir:
        output_folder = args.output_dir
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_folder = os.path.join(current_dir, f"STP体积统计结果_{timestamp}")
    os.makedirs(output_folder, exist_ok=True)

    checkpoint_path = args.checkpoint or os.path.join(output_folder, "checkpoint_progress.jsonl")
    records_map = load_checkpoint(checkpoint_path)

    pending_files = [fp for fp in stp_files if fp not in records_map]
    already_done = len(records_map)

    print(
        f"历史已完成/已隔离: {already_done} 个 | 本次待处理: {len(pending_files)} 个 | 启动 {workers} 个并行工作进程...",
        flush=True,
    )

    start_ts = time.perf_counter()
    newly_processed = 0
    abnormal_count = sum(1 for r in records_map.values() if r[6] != "成功")

    for batch_start in range(0, len(pending_files), batch_size):
        batch = pending_files[batch_start : batch_start + batch_size]
        remaining_in_batch = list(batch)

        while remaining_in_batch:
            pool_broken = False
            unfinished_files: List[str] = []

            try:
                with create_pool(workers, max_tasks_per_child) as executor:
                    future_to_file = {executor.submit(_worker_wrapper, fp): fp for fp in remaining_in_batch}
                    pending_futures = set(future_to_file.keys())

                    while pending_futures:
                        try:
                            for fut in as_completed(list(pending_futures), timeout=timeout_sec):
                                pending_futures.discard(fut)
                                fp = future_to_file[fut]
                                try:
                                    rec = fut.result()
                                except Exception:
                                    pool_broken = True
                                    unfinished_files.append(fp)
                                    continue

                                records_map[fp] = rec
                                append_checkpoint(checkpoint_path, fp, rec)
                                newly_processed += 1
                                if rec[6] != "成功":
                                    abnormal_count += 1

                                current_total = len(records_map)
                                if current_total % 100 == 0 or current_total == total_files:
                                    elapsed = time.perf_counter() - start_ts
                                    speed = newly_processed / elapsed if elapsed > 0 else 0.0
                                    pct = current_total * 100.0 / total_files
                                    print(
                                        f"进度: {current_total} / {total_files} ({pct:.1f}%) "
                                        f"| 异常/警告: {abnormal_count} | 速率: {speed:.1f} 文件/秒",
                                        flush=True,
                                    )

                                if pool_broken:
                                    break

                        except TimeoutError:
                            for fut in list(pending_futures):
                                if not fut.done():
                                    fp = future_to_file[fut]
                                    pending_futures.discard(fut)
                                    fut.cancel()
                                    fname = os.path.basename(fp)
                                    pno = os.path.splitext(fname)[0]
                                    rec = (pno, fname, fp, 0.0, 0.0, 0, f"超时跳过(>{timeout_sec:.0f}s)")
                                    records_map[fp] = rec
                                    append_checkpoint(checkpoint_path, fp, rec)
                                    newly_processed += 1
                                    abnormal_count += 1
                            pool_broken = True
                            unfinished_files.extend(
                                future_to_file[f] for f in pending_futures if not f.done()
                            )
                            break

                        if pool_broken:
                            break

                    if pool_broken:
                        executor.shutdown(wait=False, cancel_futures=True)

            except Exception as pool_exc:
                print(
                    f"\n[提示] 进程池遇到底层异常 ({type(pool_exc).__name__})，正在重建进程池并隔离疑似崩溃文件...",
                    flush=True,
                )
                pool_broken = True
                unfinished_files = [fp for fp in remaining_in_batch if fp not in records_map]

            if pool_broken:
                if unfinished_files:
                    suspect_fp = unfinished_files[0]
                    fname = os.path.basename(suspect_fp)
                    pno = os.path.splitext(fname)[0]
                    rec = (pno, fname, suspect_fp, 0.0, 0.0, 0, "崩溃隔离(底层内核异常)")
                    records_map[suspect_fp] = rec
                    append_checkpoint(checkpoint_path, suspect_fp, rec)
                    newly_processed += 1
                    abnormal_count += 1
                    print(f"[已隔离疑似崩溃文件] {suspect_fp}", flush=True)
                    remaining_in_batch = unfinished_files[1:]
                else:
                    remaining_in_batch = []
            else:
                remaining_in_batch = []

    results = [records_map[fp] for fp in stp_files if fp in records_map]
    results.sort(key=lambda x: x[0])

    batch_size_excel = 10000
    num_batches = (len(results) + batch_size_excel - 1) // batch_size_excel
    for b in range(num_batches):
        start_idx = b * batch_size_excel
        end_idx = min((b + 1) * batch_size_excel, len(results))
        batch_records = results[start_idx:end_idx]
        excel_name = f"STP体积明细_第{b + 1:02d}卷_({start_idx + 1}-{end_idx}).xlsx"
        excel_path = os.path.join(output_folder, excel_name)
        print(f"正在生成: {excel_name} ...", flush=True)
        save_to_excel(batch_records, excel_path)

    print("=" * 60, flush=True)
    print(f"全部处理完成！共统计 {len(results)} 个文件（异常/警告/超时/隔离: {abnormal_count} 个），生成 {num_batches} 个 Excel 文件。", flush=True)
    print(f"输出目录: {output_folder}", flush=True)
    print(f"断点日志: {checkpoint_path}", flush=True)
    print("=" * 60, flush=True)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    try:
        main()
    except Exception:
        traceback.print_exc()
    input("\n按回车键退出程序...")