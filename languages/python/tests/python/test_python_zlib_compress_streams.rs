#![allow(non_snake_case)]
use super::helpers::run_python;

// zlib — compress/decompress, compressobj/decompressobj flush modes, wbits, adler32, crc32

#[test]
fn test_zlib_compress_decompress_roundtrip() {
    let out = run_python(
        r#"
import zlib
data = b"hello world repeated " * 20
compressed = zlib.compress(data)
print(zlib.decompress(compressed) == data)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zlib_compress_level_0_no_compression() {
    let out = run_python(
        r#"
import zlib
data = b"a" * 1000
# level 0 = no compression, still valid zlib stream
c0 = zlib.compress(data, level=0)
c9 = zlib.compress(data, level=9)
print(len(c0) > len(c9))
print(zlib.decompress(c0) == data)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_zlib_compress_level_9_best_compression() {
    let out = run_python(
        r#"
import zlib
data = b"abcabcabc" * 100
c1 = zlib.compress(data, level=1)
c9 = zlib.compress(data, level=9)
print(len(c9) <= len(c1))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zlib_crc32_deterministic() {
    let out = run_python(
        r#"
import zlib
c1 = zlib.crc32(b"hello world")
c2 = zlib.crc32(b"hello world")
print(c1 == c2)
print(c1 != 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_zlib_crc32_different_data_differs() {
    let out = run_python(
        r#"
import zlib
print(zlib.crc32(b"aaa") != zlib.crc32(b"bbb"))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zlib_crc32_chained_value() {
    let out = run_python(
        r#"
import zlib
# crc32 of concatenated equals chained crc32
data = b"hello" + b" world"
full = zlib.crc32(data)
chained = zlib.crc32(b" world", zlib.crc32(b"hello"))
print(full == chained)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zlib_adler32_deterministic() {
    let out = run_python(
        r#"
import zlib
a1 = zlib.adler32(b"hello")
a2 = zlib.adler32(b"hello")
print(a1 == a2)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zlib_adler32_chained() {
    let out = run_python(
        r#"
import zlib
data = b"hello world"
full = zlib.adler32(data)
chained = zlib.adler32(b" world", zlib.adler32(b"hello"))
print(full == chained)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zlib_compressobj_flush_z_finish() {
    let out = run_python(
        r#"
import zlib
c = zlib.compressobj()
chunk = c.compress(b"hello world " * 50)
tail = c.flush(zlib.Z_FINISH)
compressed = chunk + tail
print(zlib.decompress(compressed) == b"hello world " * 50)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zlib_compressobj_flush_z_sync_flush() {
    let out = run_python(
        r#"
import zlib
c = zlib.compressobj()
c.compress(b"data")
synced = c.flush(zlib.Z_SYNC_FLUSH)
print(len(synced) >= 0)   # produces output, no error
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zlib_decompressobj_incremental() {
    let out = run_python(
        r#"
import zlib
data = b"hello world " * 30
compressed = zlib.compress(data)
d = zlib.decompressobj()
half = len(compressed) // 2
out1 = d.decompress(compressed[:half])
out2 = d.decompress(compressed[half:])
print(out1 + out2 == data)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zlib_compressobj_strategy_filtered() {
    let out = run_python(
        r#"
import zlib
c = zlib.compressobj(strategy=zlib.Z_FILTERED)
data = b"01010101" * 100
compressed = c.compress(data) + c.flush()
print(zlib.decompress(compressed) == data)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zlib_wbits_gzip_compress() {
    let out = run_python(
        r#"
import zlib
data = b"gzip format test " * 10
# wbits 31 = gzip format (16+15)
c = zlib.compressobj(wbits=31)
compressed = c.compress(data) + c.flush()
# decompress with wbits 47 = gzip auto-detect (16+31)
result = zlib.decompress(compressed, wbits=47)
print(result == data)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zlib_wbits_raw_deflate() {
    let out = run_python(
        r#"
import zlib
data = b"raw deflate " * 10
c = zlib.compressobj(wbits=-15)   # raw deflate, no header
compressed = c.compress(data) + c.flush()
result = zlib.decompress(compressed, wbits=-15)
print(result == data)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zlib_error_on_corrupt_data() {
    let out = run_python(
        r#"
import zlib
try:
    zlib.decompress(b"this is not valid zlib data at all XYZ")
except zlib.error:
    print("zlib.error")
"#,
    );
    assert_eq!(out, vec!["zlib.error"]);
}

#[test]
fn test_zlib_compress_empty_bytes() {
    let out = run_python(
        r#"
import zlib
compressed = zlib.compress(b"")
print(zlib.decompress(compressed) == b"")
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zlib_decompressobj_unused_data() {
    let out = run_python(
        r#"
import zlib
data = b"hello" * 10
compressed = zlib.compress(data)
trailing = b"EXTRA"
d = zlib.decompressobj()
d.decompress(compressed + trailing)
print(d.unused_data == trailing)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zlib_compressobj_mem_level() {
    let out = run_python(
        r#"
import zlib
data = b"memory level test " * 50
c = zlib.compressobj(level=6, method=zlib.DEFLATED, wbits=15, memLevel=9)
compressed = c.compress(data) + c.flush()
print(zlib.decompress(compressed) == data)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_zlib_compress_large_data() {
    let out = run_python(
        r#"
import zlib
data = bytes(range(256)) * 4000
compressed = zlib.compress(data)
print(zlib.decompress(compressed) == data)
print(len(compressed) < len(data))
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_zlib_z_best_compression_constant() {
    let out = run_python(
        r#"
import zlib
print(zlib.Z_BEST_COMPRESSION)
print(zlib.Z_BEST_SPEED)
print(zlib.Z_DEFAULT_COMPRESSION)
"#,
    );
    assert_eq!(out, vec!["9", "1", "-1"]);
}
