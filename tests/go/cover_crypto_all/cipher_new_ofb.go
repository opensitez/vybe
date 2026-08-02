// vybe-test: go/cover_crypto_all/cipher_new_ofb
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { b, _ := aes.NewCipher(make([]byte, 16))
_ = cipher.NewOFB(b, make([]byte, 16)) }
