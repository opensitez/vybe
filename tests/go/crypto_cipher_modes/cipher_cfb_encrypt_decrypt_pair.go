// vybe-test: go/crypto_cipher_modes/cipher_cfb_encrypt_decrypt_pair
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
iv := make([]byte, 16)
enc := cipher.NewCFBEncrypter(block, iv)
dec := cipher.NewCFBDecrypter(block, iv)
plain := make([]byte, 16)
ct := make([]byte, 16)
enc.XORKeyStream(ct, plain)
out := make([]byte, 16)
dec.XORKeyStream(out, ct) }
