// vybe-test: go/crypto_cipher_modes/cipher_cbc_decrypter_chaining
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
iv := make([]byte, 16)
dec := cipher.NewCBCDecrypter(block, iv)
ct := make([]byte, 32)
plain := make([]byte, 32)
dec.CryptBlocks(plain, ct) }
