"""Renders every fixture x scenario with the original Python docx-mailmerge2 (reference output).

usage: python run_python.py <fixtures_dir> <out_dir>
"""

import json
import sys
import warnings
from pathlib import Path

from mailmerge import MailMerge, MailMergeOptions, OptionKeepFields

warnings.simplefilter("ignore")
fixtures, out = Path(sys.argv[1]), Path(sys.argv[2])
scenarios = json.loads((Path(__file__).parent / "scenarios.json").read_text())
out.mkdir(parents=True, exist_ok=True)
OPTS = {"keepFields": "keep_fields", "mergeIfFields": "merge_if_fields"}


def options(o):
    kw = {OPTS[k]: v for k, v in o.items()}
    if "keep_fields" in kw:
        kw["keep_fields"] = OptionKeepFields(kw["keep_fields"])
    return MailMergeOptions(**kw)


def record(names, i, partial):
    names = sorted(names)
    if partial:
        names = names[::2]
    return {n: f"{n}#{i}" for n in names}


errors = {}
for fixture in sorted(fixtures.glob("*.docx")):
    for sc in scenarios:
        key = f"{fixture.stem}__{sc['id']}"
        try:
            with MailMerge(str(fixture), options=options(sc["options"])) as doc:
                names = doc.get_merge_fields() - {""}
                if sc["mode"] == "merge":
                    doc.merge(**record(names, 0, sc.get("partial")))
                else:
                    rows = [record(names, i, sc.get("partial")) for i in range(sc["records"])]
                    doc.merge_templates(rows, sc["separator"])
                doc.write(str(out / f"{key}.docx"))
        except Exception as e:  # noqa: BLE001
            errors[key] = f"{type(e).__name__}: {e}"
(out / "errors.json").write_text(json.dumps(errors, indent=1))
print(f"python: {len(list(out.glob('*.docx')))} documents, {len(errors)} errors")
