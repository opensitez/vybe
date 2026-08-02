// vybe-test: go/cover_crypto_all/rsa_verify_pss
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto"
import "crypto/rsa"
import "crypto/rand"
import "crypto/sha256"
func main() { key, _ := rsa.GenerateKey(rand.Reader, 512)
h := sha256.Sum256([]byte("x"))
sig, _ := rsa.SignPSS(rand.Reader, key, crypto.SHA256, h[:], nil)
_ = rsa.VerifyPSS(&key.PublicKey, crypto.SHA256, h[:], sig, nil) }
