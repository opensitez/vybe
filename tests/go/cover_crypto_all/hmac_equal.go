// vybe-test: go/cover_crypto_all/hmac_equal
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/hmac"
func main() { _ = hmac.Equal([]byte("a"), []byte("b")) }
