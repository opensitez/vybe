// vybe-test: go/crypto_cipher_modes/cipher_cbc_decrypter_roundtrip
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
iv := make([]byte, 16)
enc := cipher.NewCBCEncrypter(block, iv)
dec := cipher.NewCBCDecrypter(block, iv)
plain := make([]byte, 16)
ct := make([]byte, 16)
enc.CryptBlocks(ct, plain)
out := make([]byte, 16)
dec.CryptBlocks(out, ct) }
