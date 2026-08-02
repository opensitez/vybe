// vybe-test: go/stdlib_crypto_extended/hmac_new_sha256
// origin: languages/go/tests/go/test_stdlib_crypto_extended.rs
// vybe-test-mode: compile

package main
import "crypto/hmac"
import "crypto/sha256"
func main() { _ = hmac.New(sha256.New, []byte("key")) }
