// vybe-test: go/crypto_cipher_modes/cipher_ctr_different_ivs
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
iv1 := make([]byte, 16)
iv2 := make([]byte, 16)
iv2[15] = 1
s1 := cipher.NewCTR(block, iv1)
s2 := cipher.NewCTR(block, iv2)
_ = s1
_ = s2 }
