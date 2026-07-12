//! crypto/* beyond hash — one compile smoke per distinct API.

go_compile_cases! {
    aes_new_cipher => "package main; import \"crypto/aes\"; func main() { _, _ = aes.NewCipher(make([]byte, 16)) }",
    cipher_new_cfb_encrypter => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { b, _ := aes.NewCipher(make([]byte, 16)); _ = cipher.NewCFBEncrypter(b, make([]byte, 16)) }",
    cipher_new_gcm => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { b, _ := aes.NewCipher(make([]byte, 16)); _, _ = cipher.NewGCM(b) }",
    hmac_new_sha256 => "package main; import \"crypto/hmac\"; import \"crypto/sha256\"; func main() { _ = hmac.New(sha256.New, []byte(\"key\")) }",
    ed25519_generate_key => "package main; import \"crypto/ed25519\"; func main() { _, _, _ = ed25519.GenerateKey(nil) }",
    ed25519_sign => "package main; import \"crypto/ed25519\"; func main() { _, priv, _ := ed25519.GenerateKey(nil); _ = ed25519.Sign(priv, []byte(\"msg\")) }",
    tls_dial_compile => "package main; import \"crypto/tls\"; func main() { _, _ = tls.Dial(\"tcp\", \"example.com:443\", nil) }",
    x509_parse_certificate => "package main; import \"crypto/x509\"; func main() { _, _ = x509.ParseCertificate(nil) }",
    x509_system_cert_pool => "package main; import \"crypto/x509\"; func main() { _ = x509.SystemCertPool() }",
}
