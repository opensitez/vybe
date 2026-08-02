// vybe-test: go/cover_crypto_all/ed25519_verify
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/ed25519"
func main() { pub, priv, _ := ed25519.GenerateKey(nil)
sig := ed25519.Sign(priv, []byte("msg"))
_ = ed25519.Verify(pub, []byte("msg"), sig) }
