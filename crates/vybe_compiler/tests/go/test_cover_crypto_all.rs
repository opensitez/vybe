//! crypto/aes, des, rc4, cipher, hmac, ed25519, rsa, ecdsa, elliptic, x509, tls
//! - one distinct API per compile smoke (breadth over depth).

go_compile_cases! {
    // crypto/aes
    aes_new_cipher => "package main; import \"crypto/aes\"; func main() { _, _ = aes.NewCipher(make([]byte, 16)) }",
    aes_block_size => "package main; import \"crypto/aes\"; func main() { _ = aes.BlockSize }",

    // crypto/des
    des_new_cipher => "package main; import \"crypto/des\"; func main() { _, _ = des.NewCipher(make([]byte, 8)) }",
    des_new_triple_des_cipher => "package main; import \"crypto/des\"; func main() { _, _ = des.NewTripleDESCipher(make([]byte, 24)) }",
    des_block_size => "package main; import \"crypto/des\"; func main() { _ = des.BlockSize }",

    // crypto/rc4
    rc4_new_cipher => "package main; import \"crypto/rc4\"; func main() { _, _ = rc4.NewCipher([]byte(\"key12345\")) }",

    // crypto/cipher - block modes
    cipher_new_cbc_encrypter => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { b, _ := aes.NewCipher(make([]byte, 16)); _ = cipher.NewCBCEncrypter(b, make([]byte, 16)) }",
    cipher_new_cbc_decrypter => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { b, _ := aes.NewCipher(make([]byte, 16)); _ = cipher.NewCBCDecrypter(b, make([]byte, 16)) }",
    cipher_new_cfb_encrypter => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { b, _ := aes.NewCipher(make([]byte, 16)); _ = cipher.NewCFBEncrypter(b, make([]byte, 16)) }",
    cipher_new_cfb_decrypter => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { b, _ := aes.NewCipher(make([]byte, 16)); _ = cipher.NewCFBDecrypter(b, make([]byte, 16)) }",
    cipher_new_ctr => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { b, _ := aes.NewCipher(make([]byte, 16)); _ = cipher.NewCTR(b, make([]byte, 16)) }",
    cipher_new_gcm => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { b, _ := aes.NewCipher(make([]byte, 16)); _, _ = cipher.NewGCM(b) }",
    cipher_new_ofb => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { b, _ := aes.NewCipher(make([]byte, 16)); _ = cipher.NewOFB(b, make([]byte, 16)) }",
    cipher_subtle_constant_time_compare => "package main; import \"crypto/subtle\"; func main() { _ = subtle.ConstantTimeCompare([]byte(\"a\"), []byte(\"b\")) }",
    cipher_subtle_constant_time_copy => "package main; import \"crypto/subtle\"; func main() { dst := make([]byte, 2); subtle.ConstantTimeCopy(1, dst, []byte{1, 2}) }",
    cipher_subtle_constant_time_byte_eq => "package main; import \"crypto/subtle\"; func main() { _ = subtle.ConstantTimeByteEq(1, 2) }",
    cipher_subtle_constant_time_eq => "package main; import \"crypto/subtle\"; func main() { _ = subtle.ConstantTimeEq(1, 2) }",
    cipher_subtle_constant_time_select => "package main; import \"crypto/subtle\"; func main() { _ = subtle.ConstantTimeSelect(1, 2, 3) }",

    // crypto/hmac
    hmac_new => "package main; import \"crypto/hmac\"; import \"crypto/sha256\"; func main() { _ = hmac.New(sha256.New, []byte(\"key\")) }",
    hmac_equal => "package main; import \"crypto/hmac\"; func main() { _ = hmac.Equal([]byte(\"a\"), []byte(\"b\")) }",

    // crypto/ed25519
    ed25519_generate_key => "package main; import \"crypto/ed25519\"; func main() { _, _, _ = ed25519.GenerateKey(nil) }",
    ed25519_sign => "package main; import \"crypto/ed25519\"; func main() { _, priv, _ := ed25519.GenerateKey(nil); _ = ed25519.Sign(priv, []byte(\"msg\")) }",
    ed25519_verify => "package main; import \"crypto/ed25519\"; func main() { pub, priv, _ := ed25519.GenerateKey(nil); sig := ed25519.Sign(priv, []byte(\"msg\")); _ = ed25519.Verify(pub, []byte(\"msg\"), sig) }",
    ed25519_new_key_from_seed => "package main; import \"crypto/ed25519\"; func main() { seed := make([]byte, ed25519.SeedSize); _ = ed25519.NewKeyFromSeed(seed) }",
    ed25519_public_key_size => "package main; import \"crypto/ed25519\"; func main() { _ = ed25519.PublicKeySize }",
    ed25519_private_key_size => "package main; import \"crypto/ed25519\"; func main() { _ = ed25519.PrivateKeySize }",
    ed25519_signature_size => "package main; import \"crypto/ed25519\"; func main() { _ = ed25519.SignatureSize }",

    // crypto/rsa
    rsa_generate_key => "package main; import \"crypto/rsa\"; import \"crypto/rand\"; func main() { _, _ = rsa.GenerateKey(rand.Reader, 512) }",
    rsa_sign_pkcs1v15 => "package main; import \"crypto\"; import \"crypto/rsa\"; import \"crypto/rand\"; import \"crypto/sha256\"; func main() { key, _ := rsa.GenerateKey(rand.Reader, 512); h := sha256.Sum256([]byte(\"data\")); _, _ = rsa.SignPKCS1v15(rand.Reader, key, crypto.SHA256, h[:]) }",
    rsa_verify_pkcs1v15 => "package main; import \"crypto\"; import \"crypto/rsa\"; import \"crypto/rand\"; import \"crypto/sha256\"; func main() { key, _ := rsa.GenerateKey(rand.Reader, 512); h := sha256.Sum256([]byte(\"data\")); sig, _ := rsa.SignPKCS1v15(rand.Reader, key, crypto.SHA256, h[:]); _ = rsa.VerifyPKCS1v15(&key.PublicKey, crypto.SHA256, h[:], sig) }",
    rsa_encrypt_oaep => "package main; import \"crypto/rsa\"; import \"crypto/rand\"; import \"crypto/sha256\"; func main() { key, _ := rsa.GenerateKey(rand.Reader, 512); _, _ = rsa.EncryptOAEP(sha256.New(), rand.Reader, &key.PublicKey, []byte(\"hi\"), nil) }",
    rsa_decrypt_oaep => "package main; import \"crypto/rsa\"; import \"crypto/rand\"; import \"crypto/sha256\"; func main() { key, _ := rsa.GenerateKey(rand.Reader, 512); ct, _ := rsa.EncryptOAEP(sha256.New(), rand.Reader, &key.PublicKey, []byte(\"hi\"), nil); _, _ = rsa.DecryptOAEP(sha256.New(), rand.Reader, key, ct, nil) }",
    rsa_sign_pss => "package main; import \"crypto\"; import \"crypto/rsa\"; import \"crypto/rand\"; import \"crypto/sha256\"; func main() { key, _ := rsa.GenerateKey(rand.Reader, 512); h := sha256.Sum256([]byte(\"x\")); _, _ = rsa.SignPSS(rand.Reader, key, crypto.SHA256, h[:], nil) }",
    rsa_verify_pss => "package main; import \"crypto\"; import \"crypto/rsa\"; import \"crypto/rand\"; import \"crypto/sha256\"; func main() { key, _ := rsa.GenerateKey(rand.Reader, 512); h := sha256.Sum256([]byte(\"x\")); sig, _ := rsa.SignPSS(rand.Reader, key, crypto.SHA256, h[:], nil); _ = rsa.VerifyPSS(&key.PublicKey, crypto.SHA256, h[:], sig, nil) }",

    // crypto/ecdsa
    ecdsa_generate_key => "package main; import \"crypto/ecdsa\"; import \"crypto/elliptic\"; import \"crypto/rand\"; func main() { _, _ = ecdsa.GenerateKey(elliptic.P256(), rand.Reader) }",
    ecdsa_sign => "package main; import \"crypto/ecdsa\"; import \"crypto/elliptic\"; import \"crypto/rand\"; import \"crypto/sha256\"; func main() { key, _ := ecdsa.GenerateKey(elliptic.P256(), rand.Reader); h := sha256.Sum256([]byte(\"msg\")); _, _, _ = ecdsa.Sign(rand.Reader, key, h[:]) }",
    ecdsa_verify => "package main; import \"crypto/ecdsa\"; import \"crypto/elliptic\"; import \"crypto/rand\"; import \"crypto/sha256\"; func main() { key, _ := ecdsa.GenerateKey(elliptic.P256(), rand.Reader); h := sha256.Sum256([]byte(\"msg\")); r, s, _ := ecdsa.Sign(rand.Reader, key, h[:]); _ = ecdsa.Verify(&key.PublicKey, h[:], r, s) }",

    // crypto/elliptic
    elliptic_p224 => "package main; import \"crypto/elliptic\"; func main() { _ = elliptic.P224() }",
    elliptic_p256 => "package main; import \"crypto/elliptic\"; func main() { _ = elliptic.P256() }",
    elliptic_p384 => "package main; import \"crypto/elliptic\"; func main() { _ = elliptic.P384() }",
    elliptic_p521 => "package main; import \"crypto/elliptic\"; func main() { _ = elliptic.P521() }",
    elliptic_generate_key => "package main; import \"crypto/elliptic\"; import \"crypto/rand\"; func main() { _, _ = elliptic.GenerateKey(elliptic.P256(), rand.Reader) }",
    elliptic_marshal => "package main; import \"crypto/elliptic\"; func main() { _ = elliptic.Marshal(elliptic.P256(), []byte{1}, []byte{2}) }",
    elliptic_unmarshal => "package main; import \"crypto/elliptic\"; func main() { _, _ = elliptic.Unmarshal(elliptic.P256(), []byte{4, 1, 2}) }",

    // crypto/x509
    x509_parse_certificate => "package main; import \"crypto/x509\"; func main() { _, _ = x509.ParseCertificate(nil) }",
    x509_system_cert_pool => "package main; import \"crypto/x509\"; func main() { _ = x509.SystemCertPool() }",
    x509_parse_pkcs1_private_key => "package main; import \"crypto/x509\"; func main() { _, _ = x509.ParsePKCS1PrivateKey(nil) }",
    x509_marshal_pkcs1_private_key => "package main; import \"crypto/x509\"; import \"crypto/rsa\"; import \"crypto/rand\"; func main() { key, _ := rsa.GenerateKey(rand.Reader, 512); _ = x509.MarshalPKCS1PrivateKey(key) }",
    x509_parse_pkcs8_private_key => "package main; import \"crypto/x509\"; func main() { _, _ = x509.ParsePKCS8PrivateKey(nil) }",
    x509_marshal_pkcs8_private_key => "package main; import \"crypto/x509\"; import \"crypto/rsa\"; import \"crypto/rand\"; func main() { key, _ := rsa.GenerateKey(rand.Reader, 512); _, _ = x509.MarshalPKCS8PrivateKey(key) }",
    x509_parse_ec_private_key => "package main; import \"crypto/x509\"; func main() { _, _ = x509.ParseECPrivateKey(nil) }",
    x509_marshal_ec_private_key => "package main; import \"crypto/x509\"; import \"crypto/ecdsa\"; import \"crypto/elliptic\"; import \"crypto/rand\"; func main() { key, _ := ecdsa.GenerateKey(elliptic.P256(), rand.Reader); _, _ = x509.MarshalECPrivateKey(key) }",
    x509_parse_pkix_public_key => "package main; import \"crypto/x509\"; func main() { _, _ = x509.ParsePKIXPublicKey(nil) }",
    x509_marshal_pkix_public_key => "package main; import \"crypto/x509\"; import \"crypto/rsa\"; import \"crypto/rand\"; func main() { key, _ := rsa.GenerateKey(rand.Reader, 512); _, _ = x509.MarshalPKIXPublicKey(&key.PublicKey) }",
    x509_create_certificate => "package main; import \"crypto/x509\"; import \"crypto/rand\"; import \"crypto/rsa\"; import \"math/big\"; func main() { key, _ := rsa.GenerateKey(rand.Reader, 512); tmpl := x509.Certificate{SerialNumber: big.NewInt(1)}; _, _ = x509.CreateCertificate(rand.Reader, &tmpl, &tmpl, &key.PublicKey, key) }",
    x509_create_certificate_request => "package main; import \"crypto/x509\"; import \"crypto/rand\"; import \"crypto/rsa\"; func main() { key, _ := rsa.GenerateKey(rand.Reader, 512); _, _ = x509.CreateCertificateRequest(rand.Reader, []byte(\"csr\"), key) }",
    x509_parse_certificate_request => "package main; import \"crypto/x509\"; func main() { _, _ = x509.ParseCertificateRequest(nil) }",
    x509_parse_crl => "package main; import \"crypto/x509\"; func main() { _, _ = x509.ParseCRL(nil, nil) }",
    x509_new_cert_pool => "package main; import \"crypto/x509\"; func main() { p := x509.NewCertPool(); p.AddCert(nil) }",

    // crypto/tls
    tls_dial => "package main; import \"crypto/tls\"; func main() { _, _ = tls.Dial(\"tcp\", \"example.com:443\", nil) }",
    tls_dial_with_dialer => "package main; import \"crypto/tls\"; import \"net\"; func main() { d := net.Dialer{}; _, _ = tls.DialWithDialer(&d, \"tcp\", \"example.com:443\", nil) }",
    tls_listen => "package main; import \"crypto/tls\"; func main() { _, _ = tls.Listen(\"tcp\", \":8443\", nil) }",
    tls_new_listener => "package main; import \"crypto/tls\"; import \"net\"; func main() { ln, _ := net.Listen(\"tcp\", \":8443\"); _ = tls.NewListener(ln, nil) }",
    tls_load_x509_key_pair => "package main; import \"crypto/tls\"; func main() { _, _ = tls.LoadX509KeyPair(\"cert.pem\", \"key.pem\") }",
    tls_x509_key_pair => "package main; import \"crypto/tls\"; func main() { _, _ = tls.X509KeyPair([]byte(\"cert\"), []byte(\"key\")) }",
    tls_cipher_suites => "package main; import \"crypto/tls\"; func main() { _ = tls.CipherSuites() }",
    tls_insecure_cipher_suites => "package main; import \"crypto/tls\"; func main() { _ = tls.InsecureCipherSuites() }",
    tls_client => "package main; import \"crypto/tls\"; func main() { _ = tls.Client(nil, nil) }",
    tls_server => "package main; import \"crypto/tls\"; func main() { _ = tls.Server(nil, nil) }",
    tls_new_conn => "package main; import \"crypto/tls\"; func main() { _ = tls.NewConn(nil, nil) }",
    tls_version_name => "package main; import \"crypto/tls\"; func main() { _ = tls.VersionName(tls.VersionTLS12) }",
}
