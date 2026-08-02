// vybe-test: go/crypto_cipher_modes/cipher_new_cbc_decrypter_with_iv
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
iv := make([]byte, aes.BlockSize)
mode := cipher.NewCBCDecrypter(block, iv)
data := make([]byte, 16)
mode.CryptBlocks(data, data) }
