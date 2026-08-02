// vybe-test: go/cover_crypto_all/rc4_new_cipher
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/rc4"
func main() { _, _ = rc4.NewCipher([]byte("key12345")) }
