// vybe-test: go/cover_crypto_all/ecdsa_generate_key
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/ecdsa"
import "crypto/elliptic"
import "crypto/rand"
func main() { _, _ = ecdsa.GenerateKey(elliptic.P256(), rand.Reader) }
