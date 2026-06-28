//! Hashing: `System.Security.Cryptography` — MD5, SHA256, SHA1.
use super::helpers::run_csharp;

#[test]
fn sha256_hash_of_same_input_is_always_identical() {
    assert_eq!(
        run_csharp(
            r#"byte[] input=System.Text.Encoding.UTF8.GetBytes("hello");
var h1=System.Security.Cryptography.SHA256.HashData(input);
var h2=System.Security.Cryptography.SHA256.HashData(input);
Console.WriteLine(System.MemoryExtensions.SequenceEqual<byte>(h1,h2));"#
        ),
        &["True"]
    );
}

#[test]
fn sha256_produces_32_byte_digest() {
    assert_eq!(
        run_csharp(
            r#"var hash=System.Security.Cryptography.SHA256.HashData(new byte[]{1,2,3});
Console.WriteLine(hash.Length);"#
        ),
        &["32"]
    );
}

#[test]
fn md5_produces_16_byte_digest() {
    assert_eq!(
        run_csharp(
            r#"using var md5=System.Security.Cryptography.MD5.Create();
byte[] hash=md5.ComputeHash(new byte[]{1,2,3});
Console.WriteLine(hash.Length);"#
        ),
        &["16"]
    );
}

#[test]
fn sha1_produces_20_byte_digest() {
    assert_eq!(
        run_csharp(
            r#"var hash=System.Security.Cryptography.SHA1.HashData(new byte[]{0});
Console.WriteLine(hash.Length);"#
        ),
        &["20"]
    );
}

#[test]
fn hex_encoding_of_sha256_is_64_chars_long() {
    assert_eq!(
        run_csharp(
            r#"var hash=System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes("test"));
string hex=System.Convert.ToHexString(hash);
Console.WriteLine(hex.Length);"#
        ),
        &["64"]
    );
}
