use super::helpers::run_python;

#[test]
fn test_python_hashlib_sha256() {
    let src = r#"
import hashlib
print(hashlib.sha256(b'abc').hexdigest())
"#;
    assert_eq!(run_python(src), vec!["ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"]);
}

#[test]
fn test_python_hmac_digest() {
    let src = r#"
import hmac, hashlib
h = hmac.new(b'key', b'msg', hashlib.sha256)
print(h.hexdigest())
"#;
    assert_eq!(run_python(src), vec!["b95f8d8f6ef57d0dfb0f1f0f5f2f6cd8d3e4f9a6b4f5a8c3f4b5f9f6b6f4c8f6"]);
}
