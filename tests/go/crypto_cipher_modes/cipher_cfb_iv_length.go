// vybe-test: go/crypto_cipher_modes/cipher_cfb_iv_length
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
iv := make([]byte, block.BlockSize())
_ = cipher.NewCFBEncrypter(block, iv) }
