use super::helpers::run_python;

// codecs — encode, decode, lookup, getencoder, getdecoder, IncrementalEncoder, IncrementalDecoder, StreamWriter, StreamReader, BOM constants, register_error

#[test]
fn test_codecs_encode_decode_utf8_ascii() {
    let out = run_python(
        r#"
import codecs
b = codecs.encode("Python Codecs", "utf-8")
s = codecs.decode(b, "utf-8")
print(b)
print(s)
"#,
    );
    assert_eq!(out, vec!["b'Python Codecs'", "Python Codecs"]);
}

#[test]
fn test_codecs_lookup_codecinfo() {
    let out = run_python(
        r#"
import codecs
info = codecs.lookup("utf-8")
print(info.name)
print(callable(info.encode))
print(callable(info.decode))
"#,
    );
    assert_eq!(out, vec!["utf-8", "True", "True"]);
}

#[test]
fn test_codecs_getencoder_getdecoder() {
    let out = run_python(
        r#"
import codecs
enc = codecs.getencoder("utf-8")
dec = codecs.getdecoder("utf-8")
b, n1 = enc("hello")
s, n2 = dec(b)
print(b, n1)
print(s, n2)
"#,
    );
    assert_eq!(out, vec!["b'hello' 5", "hello 5"]);
}

#[test]
fn test_codecs_incremental_encoder_decoder_streaming() {
    let out = run_python(
        r#"
import codecs
enc = codecs.getincrementalencoder("utf-8")()
dec = codecs.getincrementaldecoder("utf-8")()

part1 = enc.encode("Hello ")
part2 = enc.encode("World!", final=True)

s1 = dec.decode(part1)
s2 = dec.decode(part2, final=True)
print(s1 + s2)
"#,
    );
    assert_eq!(out, vec!["Hello World!"]);
}

#[test]
fn test_codecs_stream_writer_and_reader() {
    let out = run_python(
        r#"
import codecs, io
buf = io.BytesIO()
writer = codecs.getwriter("utf-8")(buf)
writer.write("Streaming codecs content\n")

buf.seek(0)
reader = codecs.getreader("utf-8")(buf)
print(reader.read())
"#,
    );
    assert_eq!(out, vec!["Streaming codecs content\n"]);
}

#[test]
fn test_codecs_bom_byte_order_mark_constants() {
    let out = run_python(
        r#"
import codecs
print(codecs.BOM_UTF8)
print(codecs.BOM_UTF16_LE)
print(codecs.BOM_UTF16_BE)
"#,
    );
    assert_eq!(
        out,
        vec!["b'\\xef\\xbb\\xbf'", "b'\\xff\\xfe'", "b'\\xfe\\xff'"]
    );
}

#[test]
fn test_codecs_register_custom_error_handler() {
    let out = run_python(
        r#"
import codecs

def custom_replace(exc):
    if isinstance(exc, UnicodeDecodeError):
        return ("?", exc.end)
    raise exc

codecs.register_error("my_custom_replace", custom_replace)
res = codecs.decode(b"hello \xff world", "utf-8", errors="my_custom_replace")
print(res)
"#,
    );
    assert_eq!(out, vec!["hello ? world"]);
}

#[test]
fn test_codecs_rot13_codec_encoding() {
    let out = run_python(
        r#"
import codecs
encoded = codecs.encode("Hello World", "rot_13")
decoded = codecs.decode(encoded, "rot_13")
print(encoded)
print(decoded)
"#,
    );
    assert_eq!(out, vec!["Uryyb Jbeyq", "Hello World"]);
}

#[test]
fn test_codecs_hex_codec_transform() {
    let out = run_python(
        r#"
import codecs
hex_bytes = codecs.encode(b"abc", "hex")
raw_bytes = codecs.decode(b"616263", "hex")
print(hex_bytes)
print(raw_bytes)
"#,
    );
    assert_eq!(out, vec!["b'616263'", "b'abc'"]);
}

#[test]
fn test_codecs_base64_codec_transform() {
    let out = run_python(
        r#"
import codecs
b64 = codecs.encode(b"hello", "base64")
raw = codecs.decode(b64, "base64")
print(b64.strip())
print(raw)
"#,
    );
    assert_eq!(out, vec!["b'aGVsbG8='", "b'hello'"]);
}

#[test]
fn test_codecs_open_file_wrapping() {
    let out = run_python(
        r#"
import codecs, tempfile, os

with tempfile.NamedTemporaryFile(delete=False) as tmp:
    path = tmp.name

try:
    with codecs.open(path, "w", encoding="latin-1") as f:
        f.write("Café")

    with codecs.open(path, "r", encoding="latin-1") as f:
        content = f.read()

    print(content)
finally:
    if os.path.exists(path):
        os.remove(path)
"#,
    );
    assert_eq!(out, vec!["Café"]);
}

#[test]
fn test_codecs_ignore_error_handler() {
    let out = run_python(
        r#"
import codecs
res = codecs.decode(b"hello \xff world", "utf-8", errors="ignore")
print(res)
"#,
    );
    assert_eq!(out, vec!["hello  world"]);
}

#[test]
fn test_codecs_xmlcharrefreplace_error_handler() {
    let out = run_python(
        r#"
import codecs
res = codecs.encode("Café \u2665", "ascii", errors="xmlcharrefreplace")
print(res)
"#,
    );
    assert_eq!(out, vec!["b'Caf&#233; &#9829;'"]);
}

#[test]
fn test_codecs_namereplace_error_handler() {
    let out = run_python(
        r#"
import codecs
res = codecs.encode("\u2665", "ascii", errors="namereplace")
print(res)
"#,
    );
    assert_eq!(out, vec!["b'\\\\N{BLACK HEART SUIT}'"]);
}

#[test]
fn test_codecs_utf_16_bom_auto_detection() {
    let out = run_python(
        r#"
import codecs
encoded = codecs.BOM_UTF16_BE + "Hello".encode("utf-16-be")
decoded = codecs.decode(encoded, "utf-16")
print(decoded)
"#,
    );
    assert_eq!(out, vec!["Hello"]);
}

#[test]
fn test_codecs_lookup_unknown_encoding_raises_lookup_error() {
    let out = run_python(
        r#"
import codecs
try:
    codecs.lookup("non_existent_codec_999")
except LookupError:
    print("LookupError")
"#,
    );
    assert_eq!(out, vec!["LookupError"]);
}

#[test]
fn test_codecs_encoded_file_wrapping() {
    let out = run_python(
        r#"
import codecs, io
buf = io.BytesIO()
ef = codecs.EncodedFile(buf, data_encoding="utf-8", file_encoding="utf-8")
ef.write(b"encoded file data")
buf.seek(0)
print(buf.read())
"#,
    );
    assert_eq!(out, vec!["b'encoded file data'"]);
}

#[test]
fn test_codecs_escape_decode_and_encode() {
    let out = run_python(
        r#"
import codecs
raw = r"hello\nworld"
decoded, _ = codecs.escape_decode(raw.encode("ascii"))
print(decoded)
"#,
    );
    assert_eq!(out, vec!["b'hello\\nworld'"]);
}

#[test]
fn test_codecs_iterencode_generator() {
    let out = run_python(
        r#"
import codecs
chunks = ["Hello ", "World"]
encoded_chunks = list(codecs.iterencode(chunks, "utf-8"))
print(b"".join(encoded_chunks))
"#,
    );
    assert_eq!(out, vec!["b'Hello World'"]);
}

#[test]
fn test_codecs_iterdecode_generator() {
    let out = run_python(
        r#"
import codecs
chunks = [b"Hello ", b"World"]
decoded_chunks = list(codecs.iterdecode(chunks, "utf-8"))
print("".join(decoded_chunks))
"#,
    );
    assert_eq!(out, vec!["Hello World"]);
}
