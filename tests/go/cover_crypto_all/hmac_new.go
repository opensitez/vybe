// vybe-test: go/cover_crypto_all/hmac_new
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/hmac"
import "crypto/sha256"
func main() { _ = hmac.New(sha256.New, []byte("key")) }
