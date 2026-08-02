// vybe-test: go/stdlib_crypto_extended/aes_new_cipher
// origin: languages/go/tests/go/test_stdlib_crypto_extended.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
func main() { _, _ = aes.NewCipher(make([]byte, 16)) }
