// vybe-test: go/cover_crypto_all/ed25519_new_key_from_seed
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/ed25519"
func main() { seed := make([]byte, ed25519.SeedSize)
_ = ed25519.NewKeyFromSeed(seed) }
