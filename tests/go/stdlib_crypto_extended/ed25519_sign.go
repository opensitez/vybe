// vybe-test: go/stdlib_crypto_extended/ed25519_sign
// origin: languages/go/tests/go/test_stdlib_crypto_extended.rs
// vybe-test-mode: compile

package main
import "crypto/ed25519"
func main() { _, priv, _ := ed25519.GenerateKey(nil)
_ = ed25519.Sign(priv, []byte("msg")) }
