//! crypto/aes cipher modes: block size, ECB-style Encrypt/Decrypt, CBC IV, GCM seal/open,
//! crypto/cipher Stream — distinct from `test_crypto_hash_compile.rs` and compile smokes in
//! `test_stdlib_crypto_extended.rs` / `test_cover_crypto_all.rs`.

use crate::helpers::*;

go_run_cases! {
    aes_block_size_constant => (
        "package main; import \"fmt\"; import \"crypto/aes\"; func main() { fmt.Println(aes.BlockSize) }",
        vec!["16"]
    ),
    aes_new_cipher_valid_key_len => (
        "package main; import \"fmt\"; import \"crypto/aes\"; func main() { _, err := aes.NewCipher(make([]byte, 16)); fmt.Println(err == nil) }",
        vec!["true"]
    ),
    aes_new_cipher_24_byte_key => (
        "package main; import \"fmt\"; import \"crypto/aes\"; func main() { _, err := aes.NewCipher(make([]byte, 24)); fmt.Println(err == nil) }",
        vec!["true"]
    ),
    aes_new_cipher_32_byte_key => (
        "package main; import \"fmt\"; import \"crypto/aes\"; func main() { _, err := aes.NewCipher(make([]byte, 32)); fmt.Println(err == nil) }",
        vec!["true"]
    ),
    aes_new_cipher_invalid_key_len => (
        "package main; import \"fmt\"; import \"crypto/aes\"; func main() { _, err := aes.NewCipher(make([]byte, 15)); fmt.Println(err != nil) }",
        vec!["true"]
    ),
    aes_encrypt_decrypt_single_block => (
        "package main; import \"fmt\"; import \"crypto/aes\"; func main() { key := make([]byte, 16); block, _ := aes.NewCipher(key); plain := make([]byte, 16); cipher := make([]byte, 16); block.Encrypt(cipher, plain); out := make([]byte, 16); block.Decrypt(out, cipher); fmt.Println(out[0] == plain[0]) }",
        vec!["true"]
    ),
    aes_encrypt_changes_ciphertext => (
        "package main; import \"fmt\"; import \"crypto/aes\"; func main() { key := []byte{1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16}; block, _ := aes.NewCipher(key); plain := make([]byte, 16); for i := range plain { plain[i] = byte(i) }; cipher := make([]byte, 16); block.Encrypt(cipher, plain); fmt.Println(cipher[0] != plain[0]) }",
        vec!["true"]
    ),
    aes_block_size_matches_buffer => (
        "package main; import \"fmt\"; import \"crypto/aes\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); fmt.Println(block.BlockSize()) }",
        vec!["16"]
    ),
    aes_encrypt_two_blocks_ecb_pattern => (
        "package main; import \"fmt\"; import \"crypto/aes\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); p1 := make([]byte, 16); p2 := make([]byte, 16); p2[0] = 1; c1 := make([]byte, 16); c2 := make([]byte, 16); block.Encrypt(c1, p1); block.Encrypt(c2, p2); fmt.Println(c1[0] != c2[0]) }",
        vec!["true"]
    ),
    aes_decrypt_restores_plaintext => (
        "package main; import \"fmt\"; import \"crypto/aes\"; func main() { block, _ := aes.NewCipher([]byte{0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15}); plain := []byte{0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1}; cipher := make([]byte, 16); block.Encrypt(cipher, plain); out := make([]byte, 16); block.Decrypt(out, cipher); fmt.Println(out[15]) }",
        vec!["1"]
    ),
}

go_compile_cases! {
    aes_ecb_encrypt_loop_two_blocks => "package main; import \"crypto/aes\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); plain := make([]byte, 32); cipher := make([]byte, 32); for i := 0; i < len(plain); i += block.BlockSize() { block.Encrypt(cipher[i:i+16], plain[i:i+16]) } }",
    aes_ecb_decrypt_loop_two_blocks => "package main; import \"crypto/aes\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); cipher := make([]byte, 32); plain := make([]byte, 32); for i := 0; i < len(cipher); i += block.BlockSize() { block.Decrypt(plain[i:i+16], cipher[i:i+16]) } }",
    cipher_new_cbc_encrypter_with_iv => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); iv := make([]byte, aes.BlockSize); mode := cipher.NewCBCEncrypter(block, iv); data := make([]byte, 16); mode.CryptBlocks(data, data) }",
    cipher_new_cbc_decrypter_with_iv => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); iv := make([]byte, aes.BlockSize); mode := cipher.NewCBCDecrypter(block, iv); data := make([]byte, 16); mode.CryptBlocks(data, data) }",
    cipher_cbc_encrypter_two_blocks => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); iv := make([]byte, 16); enc := cipher.NewCBCEncrypter(block, iv); plain := make([]byte, 32); cipherText := make([]byte, 32); enc.CryptBlocks(cipherText, plain) }",
    cipher_cbc_decrypter_roundtrip => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); iv := make([]byte, 16); enc := cipher.NewCBCEncrypter(block, iv); dec := cipher.NewCBCDecrypter(block, iv); plain := make([]byte, 16); ct := make([]byte, 16); enc.CryptBlocks(ct, plain); out := make([]byte, 16); dec.CryptBlocks(out, ct) }",
    cipher_cbc_different_iv => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); iv1 := make([]byte, 16); iv2 := make([]byte, 16); iv2[0] = 1; _ = cipher.NewCBCEncrypter(block, iv1); _ = cipher.NewCBCEncrypter(block, iv2) }",
    cipher_new_gcm_seal => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; import \"crypto/rand\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); gcm, _ := cipher.NewGCM(block); nonce := make([]byte, gcm.NonceSize()); plain := []byte(\"secret\"); _ = gcm.Seal(nil, nonce, plain, nil) }",
    cipher_new_gcm_open => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); gcm, _ := cipher.NewGCM(block); nonce := make([]byte, gcm.NonceSize()); ct := gcm.Seal(nil, nonce, []byte(\"data\"), nil); _, _ = gcm.Open(nil, nonce, ct, nil) }",
    cipher_gcm_with_aad => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); gcm, _ := cipher.NewGCM(block); nonce := make([]byte, gcm.NonceSize()); aad := []byte(\"header\"); _ = gcm.Seal(nil, nonce, []byte(\"body\"), aad) }",
    cipher_gcm_nonce_size => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); gcm, _ := cipher.NewGCM(block); _ = gcm.NonceSize() }",
    cipher_gcm_overhead => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); gcm, _ := cipher.NewGCM(block); _ = gcm.Overhead() }",
    cipher_new_gcm_with_tag_size => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); _, _ = cipher.NewGCMWithTagSize(block, 12) }",
    cipher_new_gcm_with_nonce_size => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); _, _ = cipher.NewGCMWithNonceSize(block, 12) }",
    cipher_new_ctr_stream => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); stream := cipher.NewCTR(block, make([]byte, 16)); dst := make([]byte, 16); stream.XORKeyStream(dst, dst) }",
    cipher_new_cfb_encrypter_stream => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); stream := cipher.NewCFBEncrypter(block, make([]byte, 16)); dst := make([]byte, 16); stream.XORKeyStream(dst, dst) }",
    cipher_new_cfb_decrypter_stream => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); stream := cipher.NewCFBDecrypter(block, make([]byte, 16)); dst := make([]byte, 16); stream.XORKeyStream(dst, dst) }",
    cipher_new_ofb_stream => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); stream := cipher.NewOFB(block, make([]byte, 16)); dst := make([]byte, 32); stream.XORKeyStream(dst, dst) }",
    cipher_stream_interface_xor => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); var s cipher.Stream = cipher.NewCTR(block, make([]byte, 16)); buf := make([]byte, 8); s.XORKeyStream(buf, buf) }",
    cipher_stream_reader => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; import \"io\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); stream := cipher.NewCTR(block, make([]byte, 16)); r := &cipher.StreamReader{S: stream, R: nil}; _ = r }",
    cipher_stream_writer => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; import \"bytes\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); stream := cipher.NewCTR(block, make([]byte, 16)); w := &cipher.StreamWriter{S: stream, W: bytes.NewBuffer(nil)}; _, _ = w.Write([]byte(\"x\")) }",
    aes_encrypt_dst_src_overlap => "package main; import \"crypto/aes\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); buf := make([]byte, 16); block.Encrypt(buf, buf) }",
    aes_decrypt_dst_src_overlap => "package main; import \"crypto/aes\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); buf := make([]byte, 16); block.Decrypt(buf, buf) }",
    cipher_cbc_iv_must_match_block => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); iv := make([]byte, aes.BlockSize); _ = cipher.NewCBCEncrypter(block, iv) }",
    cipher_gcm_seal_dst_preallocated => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); gcm, _ := cipher.NewGCM(block); nonce := make([]byte, gcm.NonceSize()); dst := make([]byte, 0, 64); _ = gcm.Seal(dst, nonce, []byte(\"hi\"), nil) }",
    cipher_gcm_open_into_dst => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); gcm, _ := cipher.NewGCM(block); nonce := make([]byte, gcm.NonceSize()); sealed := gcm.Seal(nil, nonce, []byte(\"msg\"), nil); _, _ = gcm.Open(sealed[:0], nonce, sealed, nil) }",
    cipher_ctr_different_ivs => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); iv1 := make([]byte, 16); iv2 := make([]byte, 16); iv2[15] = 1; s1 := cipher.NewCTR(block, iv1); s2 := cipher.NewCTR(block, iv2); _ = s1; _ = s2 }",
    cipher_cfb_encrypt_decrypt_pair => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); iv := make([]byte, 16); enc := cipher.NewCFBEncrypter(block, iv); dec := cipher.NewCFBDecrypter(block, iv); plain := make([]byte, 16); ct := make([]byte, 16); enc.XORKeyStream(ct, plain); out := make([]byte, 16); dec.XORKeyStream(out, ct) }",
    aes_new_cipher_copy_key => "package main; import \"crypto/aes\"; func main() { key := []byte{0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15}; block, _ := aes.NewCipher(key); key[0] = 99; plain := make([]byte, 16); cipher := make([]byte, 16); block.Encrypt(cipher, plain) }",
    cipher_block_interface => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { var b cipher.Block; b, _ = aes.NewCipher(make([]byte, 16)); _ = b.BlockSize() }",
    cipher_cbc_encrypt_chaining => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); iv := make([]byte, 16); enc := cipher.NewCBCEncrypter(block, iv); p := make([]byte, 32); p[16] = 1; c := make([]byte, 32); enc.CryptBlocks(c, p) }",
    cipher_gcm_open_wrong_nonce => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); gcm, _ := cipher.NewGCM(block); nonce := make([]byte, gcm.NonceSize()); sealed := gcm.Seal(nil, nonce, []byte(\"x\"), nil); badNonce := make([]byte, gcm.NonceSize()); badNonce[0] = 1; _, err := gcm.Open(nil, badNonce, sealed, nil); _ = err }",
    cipher_stream_xor_partial => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); s := cipher.NewCTR(block, make([]byte, 16)); dst := make([]byte, 5); src := []byte(\"hello\"); s.XORKeyStream(dst, src) }",
    aes_key_schedule_128 => "package main; import \"crypto/aes\"; func main() { block, _ := aes.NewCipher([]byte{0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15}); _ = block }",
    aes_key_schedule_192 => "package main; import \"crypto/aes\"; func main() { block, _ := aes.NewCipher(make([]byte, 24)); _ = block }",
    aes_key_schedule_256 => "package main; import \"crypto/aes\"; func main() { block, _ := aes.NewCipher(make([]byte, 32)); _ = block }",
    cipher_ofb_stream_keystream => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); s := cipher.NewOFB(block, make([]byte, 16)); buf := make([]byte, 64); s.XORKeyStream(buf, buf) }",
    cipher_cbc_decrypter_chaining => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); iv := make([]byte, 16); dec := cipher.NewCBCDecrypter(block, iv); ct := make([]byte, 32); plain := make([]byte, 32); dec.CryptBlocks(plain, ct) }",
    cipher_gcm_seal_empty_plaintext => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); gcm, _ := cipher.NewGCM(block); nonce := make([]byte, gcm.NonceSize()); _ = gcm.Seal(nil, nonce, []byte{}, nil) }",
    cipher_stream_reader_read => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; import \"bytes\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); s := cipher.NewCTR(block, make([]byte, 16)); r := &cipher.StreamReader{S: s, R: bytes.NewReader([]byte(\"data\"))}; buf := make([]byte, 4); _, _ = r.Read(buf) }",
    cipher_stream_writer_close => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; import \"bytes\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); s := cipher.NewCTR(block, make([]byte, 16)); w := &cipher.StreamWriter{S: s, W: bytes.NewBuffer(nil)}; _ = w.Close() }",
    cipher_new_gcm_256_bit_key => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 32)); gcm, _ := cipher.NewGCM(block); _ = gcm.NonceSize() }",
    cipher_cbc_encrypter_block_size => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); iv := make([]byte, 16); enc := cipher.NewCBCEncrypter(block, iv); _ = enc }",
    cipher_ctr_iv_length => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); iv := make([]byte, block.BlockSize()); _ = cipher.NewCTR(block, iv) }",
    cipher_cfb_iv_length => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); iv := make([]byte, block.BlockSize()); _ = cipher.NewCFBEncrypter(block, iv) }",
    cipher_ofb_iv_length => "package main; import \"crypto/aes\"; import \"crypto/cipher\"; func main() { block, _ := aes.NewCipher(make([]byte, 16)); iv := make([]byte, block.BlockSize()); _ = cipher.NewOFB(block, iv) }",
}
