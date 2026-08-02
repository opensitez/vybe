// vybe-test: go/crypto_cipher_modes/cipher_cbc_different_iv
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
iv1 := make([]byte, 16)
iv2 := make([]byte, 16)
iv2[0] = 1
_ = cipher.NewCBCEncrypter(block, iv1)
_ = cipher.NewCBCEncrypter(block, iv2) }
