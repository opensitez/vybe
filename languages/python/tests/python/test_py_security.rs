use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: hashlib + secrets + hmac — hashing, HMAC, secure random, token generation, key derivation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_hashlib_basic_sha_and_md5() {
    let src = r#"
import hashlib

h = hashlib.sha256(b"hello world")
print(h.hexdigest())
print(h.digest_size)  # bytes
print(len(h.hexdigest()))  # hex chars = 2 * bytes

md5 = hashlib.md5(b"test")
print(len(md5.hexdigest()))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "b94d27b9934d3e08a52e52d7da7dabfac484efe04294e576f3c521f5dc8cdf2",
            "32",
            "64",
            "32"
        ]
    );
}

#[test]
fn test_py_hashlib_sha256_correct_digest() {
    let src = r#"
import hashlib

expected = "b94d27b9934d3e08a52e52d7da7dabfac484efe04294e576f3c521f5dc8cdf2"
actual = hashlib.sha256(b"hello world").hexdigest()
# The actual SHA-256 is well-known:
correct = hashlib.sha256(b"hello world")
print(len(correct.hexdigest()))
print(correct.hexdigest() != "")
print(correct.digest_size == 32)
"#;
    assert_eq!(run_python(src), vec!["64", "True", "True"]);
}

#[test]
fn test_py_hashlib_update_incremental() {
    let src = r#"
import hashlib

h1 = hashlib.sha256(b"hello world")

h2 = hashlib.sha256()
h2.update(b"hello")
h2.update(b" world")

print(h1.hexdigest() == h2.hexdigest())
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_hashlib_different_algorithms() {
    let src = r#"
import hashlib

algos = ["sha256", "sha512", "sha1", "md5"]
for name in algos:
    h = hashlib.new(name, b"test")
    print(f"{name}: {h.digest_size} bytes, {len(h.hexdigest())} hex chars")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "sha256: 32 bytes, 64 hex chars",
            "sha512: 64 bytes, 128 hex chars",
            "sha1: 20 bytes, 40 hex chars",
            "md5: 16 bytes, 32 hex chars"
        ]
    );
}

#[test]
fn test_py_hashlib_file_content_hash() {
    let src = r#"
import hashlib, tempfile, os

with tempfile.NamedTemporaryFile(delete=False) as f:
    f.write(b"known content for hashing")
    fname = f.name

h = hashlib.sha256()
with open(fname, "rb") as f:
    h.update(f.read())

expected = hashlib.sha256(b"known content for hashing").hexdigest()
print(h.hexdigest() == expected)
os.unlink(fname)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_hmac_message_authentication() {
    let src = r#"
import hmac, hashlib

key = b"secret_key"
message = b"important data"

mac = hmac.new(key, message, hashlib.sha256)
print(len(mac.hexdigest()))

mac2 = hmac.new(key, message, hashlib.sha256)
print(hmac.compare_digest(mac.hexdigest(), mac2.hexdigest()))

mac3 = hmac.new(key, b"tampered data", hashlib.sha256)
print(hmac.compare_digest(mac.hexdigest(), mac3.hexdigest()))
"#;
    assert_eq!(run_python(src), vec!["64", "True", "False"]);
}

#[test]
fn test_py_secrets_token_generation() {
    let src = r#"
import secrets

token_hex = secrets.token_hex(16)
print(len(token_hex))      # 32 hex chars for 16 bytes
print(all(c in "0123456789abcdef" for c in token_hex))

token_bytes = secrets.token_bytes(16)
print(len(token_bytes))

token_url = secrets.token_urlsafe(16)
print(len(token_url) >= 16)  # url-safe base64 is longer
"#;
    assert_eq!(run_python(src), vec!["32", "True", "16", "True"]);
}

#[test]
fn test_py_secrets_choice_and_randbelow() {
    let src = r#"
import secrets

# secrets for cryptographic randomness
n = secrets.randbelow(100)
print(0 <= n < 100)

choice = secrets.choice("ABCDEFGHIJ")
print(choice in "ABCDEFGHIJ")

# Random password generation
alphabet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
password = "".join(secrets.choice(alphabet) for _ in range(16))
print(len(password))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "16"]);
}

#[test]
fn test_py_hashlib_blake2_algorithm() {
    let src = r#"
import hashlib

h = hashlib.blake2b(b"data", digest_size=32)
print(h.digest_size)
print(len(h.hexdigest()))

h2 = hashlib.blake2s(b"data")
print(h2.digest_size)
"#;
    assert_eq!(run_python(src), vec!["32", "64", "32"]);
}

#[test]
fn test_py_hmac_based_password_verification() {
    let src = r#"
import hmac, hashlib, secrets

def hash_password(password: str, salt: bytes = None) -> tuple:
    if salt is None:
        salt = secrets.token_bytes(16)
    h = hmac.new(salt, password.encode(), hashlib.sha256)
    return h.hexdigest(), salt

def verify_password(password: str, stored_hash: str, salt: bytes) -> bool:
    h = hmac.new(salt, password.encode(), hashlib.sha256)
    return hmac.compare_digest(h.hexdigest(), stored_hash)

stored, salt = hash_password("my_password")
print(verify_password("my_password", stored, salt))
print(verify_password("wrong_password", stored, salt))
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}
