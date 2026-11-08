"""作者：Wangyuan；日期：2026-10-29。完全合成的离线库调用演示。"""
import argparse,json
argparse.ArgumentParser(description="纯合成数据演示，不联网，不读取原业务目录").parse_args()
import tempfile
from pathlib import Path
from dataclasses import asdict
from openpyxl import Workbook,load_workbook
from offline_office_core.compare import TableSchema,compare_files,write_changes
with tempfile.TemporaryDirectory() as folder:
    root=Path(folder)
    for name,end in [("old","00:30"),("new","00:45")]:
        wb=Workbook();wb.active.append(["id","start","end"]);wb.active.append(["synthetic","23:30",end]);wb.save(root/f"{name}.xlsx")
    schema=TableSchema(("id",),("start","end"),"start","end")
    changes=compare_files(root/"old.xlsx",root/"new.xlsx",schema)
    output=write_changes(changes,root/"changes.xlsx");readback=load_workbook(output)
    print(json.dumps({"synthetic":True,"changes":[asdict(x) for x in changes],"output_rows":readback.active.max_row},ensure_ascii=False,indent=2));readback.close()
