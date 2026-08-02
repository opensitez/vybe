// vybe-test: go/cover_crypto_all/rsa_sign_pss
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto"
import "crypto/rsa"
import "crypto/rand"
import "crypto/sha256"
func main() { key, _ := rsa.GenerateKey(rand.Reader, 512)
h := sha256.Sum256([]byte("x"))
_, _ = rsa.SignPSS(rand.Reader, key, crypto.SHA256, h[:], nil) }
