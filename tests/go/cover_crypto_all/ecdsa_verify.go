// vybe-test: go/cover_crypto_all/ecdsa_verify
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/ecdsa"
import "crypto/elliptic"
import "crypto/rand"
import "crypto/sha256"
func main() { key, _ := ecdsa.GenerateKey(elliptic.P256(), rand.Reader)
h := sha256.Sum256([]byte("msg"))
r, s, _ := ecdsa.Sign(rand.Reader, key, h[:])
_ = ecdsa.Verify(&key.PublicKey, h[:], r, s) }
