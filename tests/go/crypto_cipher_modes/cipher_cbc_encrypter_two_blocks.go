// vybe-test: go/crypto_cipher_modes/cipher_cbc_encrypter_two_blocks
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
iv := make([]byte, 16)
enc := cipher.NewCBCEncrypter(block, iv)
plain := make([]byte, 32)
cipherText := make([]byte, 32)
enc.CryptBlocks(cipherText, plain) }
