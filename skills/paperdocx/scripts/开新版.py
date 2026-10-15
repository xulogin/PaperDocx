# -*- coding: utf-8 -*-
r"""开一个新版本：把当前最新的三件套复制成 V(N+1)，旧版原样冻住。

    py -3.12 -X utf8 _工作区/排版/开新版.py            # V41 → V42
    py -3.12 -X utf8 _工作区/排版/开新版.py --从 V41 --到 V42

只复制两份 md（docx 由 `出Word.py` 重新出，不复制 —— 新版的书签号会变）。
已存在就拒绝覆盖。

开完新版的标准动作：

    ① 在 C_英文稿V42_英文版.md 上改英文
    ② 中文大概跟一下（改哪块跟哪块，铁律 3）
    ③ py -3.12 -X utf8 _工作区/排版/出Word.py --标红对照 C_英文稿V41_英文版.md
       （和上一版 diff，改过的地方在 Word 里标红）
    ④ py -3.12 -X utf8 _工作区/排版/三方同步.py --重置快照 --写
    ⑤ py -3.12 -X utf8 _工作区/自检.py
    ⑥ 冻基线 + 写 进展.md
"""
import re
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _项目                                           # noqa: E402
根 = _项目.找根()


def 取参(名, 缺省=None):
    return sys.argv[sys.argv.index(名) + 1] if 名 in sys.argv else 缺省


def 主():
    从 = 取参("--从")
    if not 从:
        从 = _项目.当前版(根)
    到 = 取参("--到") or re.sub(r"\d+", lambda m: str(int(m.group()) + 1), 从, count=1)

    c = _项目.配置(根)
    对 = []
    for 键 in ("英文版名", "对照版名"):
        源 = 根 / c[键].format(版=从)
        目 = 根 / c[键].format(版=到)
        if not 源.exists():
            raise SystemExit(f"⛔ 找不到 {源.name}")
        if 目.exists():
            raise SystemExit(f"⛔ {目.name} 已存在，不覆盖。真要重开先自己挪走。")
        对.append((源, 目))
    for 源, 目 in 对:
        shutil.copyfile(源, 目)
        print(f"✓ {源.name} → {目.name}  ({目.stat().st_size} 字节)")

    print(f"\n{从} 的三件套原样冻住，别再动它。接下来在 C_英文稿{到}_英文版.md 上改，")
    print(f"改完跑：")
    print(f"  py -3.12 -X utf8 _工作区/排版/出Word.py --标红对照 C_英文稿{从}_英文版.md")
    print(f"  py -3.12 -X utf8 _工作区/排版/三方同步.py --重置快照 --写")
    print(f"  py -3.12 -X utf8 _工作区/自检.py")


if __name__ == "__main__":
    主()
