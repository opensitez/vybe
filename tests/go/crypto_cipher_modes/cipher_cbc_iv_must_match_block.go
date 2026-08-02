// vybe-test: go/crypto_cipher_modes/cipher_cbc_iv_must_match_block
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
iv := make([]byte, aes.BlockSize)
_ = cipher.NewCBCEncrypter(block, iv) }
