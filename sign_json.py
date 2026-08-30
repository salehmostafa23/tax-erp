import base64, datetime, json, hashlib, os, sys

APP_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")

def load_sign_lib():
    src = open(APP_FILE, encoding="utf-8").read()
    start = src.index("def eta_strip_nulls")
    end = src.index("def _extract_barcode_from_line")
    block = src[start:end]
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.serialization import pkcs12, Encoding
    from cryptography.hazmat.primitives.asymmetric import padding
    ns = dict(base64=base64, datetime=datetime.datetime, json=json, hashlib=hashlib, os=os,
              hashes=hashes, pkcs12=pkcs12, Encoding=Encoding, padding=padding, serialization=serialization,
              __file__=APP_FILE)
    exec(compile(block, "<eta_sign_lib>", "exec"), ns)
    return ns

def main():
    if len(sys.argv) < 2:
        print("Usage: python sign_json.py <input.json> [output.json]")
        sys.exit(2)
    inp = sys.argv[1]
    if not os.path.exists(inp):
        print("file not found:", inp); sys.exit(2)
    out = sys.argv[2] if len(sys.argv) > 2 else (inp[:-5] + ".signed.json" if inp.lower().endswith(".json") else inp + ".signed.json")
    doc = json.load(open(inp, encoding="utf-8"))
    if "signatures" in doc:
        doc["signatures"] = []
    lib = load_sign_lib()
    res = lib["eta_sign_json_document"](doc, None, "")
    if res.get("error"):
        print("signing failed:", res["error"]); sys.exit(1)
    if not res.get("signature"):
        print("no eSeal smartcard on this machine - plug the token and run the local app (or upload a PFX).")
        sys.exit(1)
    with open(out, "w", encoding="utf-8") as f:
        json.dump(res["document"], f, ensure_ascii=False, indent=2)
    print("signed successfully. output:", out)

if __name__ == "__main__":
    main()