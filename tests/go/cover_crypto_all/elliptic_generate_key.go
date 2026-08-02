// vybe-test: go/cover_crypto_all/elliptic_generate_key
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/elliptic"
import "crypto/rand"
func main() { _, _ = elliptic.GenerateKey(elliptic.P256(), rand.Reader) }
