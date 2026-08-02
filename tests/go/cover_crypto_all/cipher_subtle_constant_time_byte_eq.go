// vybe-test: go/cover_crypto_all/cipher_subtle_constant_time_byte_eq
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/subtle"
func main() { _ = subtle.ConstantTimeByteEq(1, 2) }
