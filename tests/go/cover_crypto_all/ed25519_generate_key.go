// vybe-test: go/cover_crypto_all/ed25519_generate_key
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/ed25519"
func main() { _, _, _ = ed25519.GenerateKey(nil) }
