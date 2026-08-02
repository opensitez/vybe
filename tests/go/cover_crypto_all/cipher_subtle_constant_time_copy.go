// vybe-test: go/cover_crypto_all/cipher_subtle_constant_time_copy
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/subtle"
func main() { dst := make([]byte, 2)
subtle.ConstantTimeCopy(1, dst, []byte{1, 2}) }
