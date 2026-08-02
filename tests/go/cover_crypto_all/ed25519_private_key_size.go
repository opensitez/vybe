// vybe-test: go/cover_crypto_all/ed25519_private_key_size
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/ed25519"
func main() { _ = ed25519.PrivateKeySize }
