// vybe-test: go/cover_crypto_all/aes_new_cipher
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
func main() { _, _ = aes.NewCipher(make([]byte, 16)) }
