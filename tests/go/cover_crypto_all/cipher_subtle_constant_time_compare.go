// vybe-test: go/cover_crypto_all/cipher_subtle_constant_time_compare
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/subtle"
func main() { _ = subtle.ConstantTimeCompare([]byte("a"), []byte("b")) }
