"""Exact audit FNV fingerprints using a tiny, archived native helper."""
import ctypes
import hashlib
from pathlib import Path
import subprocess
import sys


def build(directory: Path, python_reference):
    directory.mkdir(parents=True, exist_ok=True)
    source = directory / "fingerprint.c"
    source.write_text('#include <stdint.h>\n#include <stddef.h>\nuint64_t feed(uint64_t h, const unsigned char *p, size_t n) { while(n--) h=(h^*p++)*UINT64_C(0x100000001b3); return h; }\n')
    library = directory / "fingerprint.so"
    command = ["cc", "-O3", "-dynamiclib" if sys.platform == "darwin" else "-shared", "-fPIC", str(source), "-o", str(library)]
    subprocess.run(command, check=True, capture_output=True)
    feed = ctypes.CDLL(str(library)).feed
    feed.argtypes = [ctypes.c_uint64, ctypes.c_char_p, ctypes.c_size_t]
    feed.restype = ctypes.c_uint64

    def fingerprint(label, data):
        value = 0xcbf29ce484222325
        for part in (label.encode(), data):
            size = len(part).to_bytes(8, "little")
            value = feed(value, size, 8)
            value = feed(value, part, len(part))
        return f"{value:016x}"

    for label, data in (("", b""), ("probe", b"\x00abc"), ("é\t\n\\", bytes(range(256))), ("long", b"abc\x00" * 4096)):
        assert fingerprint(label, data) == python_reference(label, data)
    return fingerprint, {"command": command, "source_sha256": hashlib.sha256(source.read_bytes()).hexdigest(), "library_sha256": hashlib.sha256(library.read_bytes()).hexdigest()}
