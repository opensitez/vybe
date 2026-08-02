// vybe-test: go/stdlib_crypto_extended/cipher_new_gcm
// origin: languages/go/tests/go/test_stdlib_crypto_extended.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { b, _ := aes.NewCipher(make([]byte, 16))
_, _ = cipher.NewGCM(b) }
