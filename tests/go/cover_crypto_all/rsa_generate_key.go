// vybe-test: go/cover_crypto_all/rsa_generate_key
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/rsa"
import "crypto/rand"
func main() { _, _ = rsa.GenerateKey(rand.Reader, 512) }
