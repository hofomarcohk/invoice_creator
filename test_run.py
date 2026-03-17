"""Headless test — no GUI needed."""
import sys, os, traceback

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

import invoice_generator as ig
ig.resource_path = lambda p: os.path.join(BASE, p)

output_dir = os.path.join(BASE, "test_output")
os.makedirs(output_dir, exist_ok=True)

log = []
def logf(msg):
    print(msg, flush=True)
    log.append(msg)

try:
    invoice_data, bill_data = ig.read_import(os.path.join(BASE, "import_sample.xlsx"))
    logf(f"Clients found: {list(invoice_data.keys())}")

    for code, info in invoice_data.items():
        txns = bill_data.get(code, [])
        logf(f"\n[{code}] {len(txns)} transaction(s)")
        ig.generate_client_files(code, info, txns, output_dir, "2026/03/16", logf)

    logf("\nSUCCESS")
except Exception as e:
    logf(f"\nERROR: {e}")
    logf(traceback.format_exc())

with open(os.path.join(BASE, "test_run.log"), "w", encoding="utf-8") as f:
    f.write("\n".join(log))
