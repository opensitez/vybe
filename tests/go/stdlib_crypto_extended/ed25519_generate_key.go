// vybe-test: go/stdlib_crypto_extended/ed25519_generate_key
// origin: languages/go/tests/go/test_stdlib_crypto_extended.rs
// vybe-test-mode: compile

package main
import "crypto/ed25519"
func main() { _, _, _ = ed25519.GenerateKey(nil) }
