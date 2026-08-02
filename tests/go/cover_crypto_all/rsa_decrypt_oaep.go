// vybe-test: go/cover_crypto_all/rsa_decrypt_oaep
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/rsa"
import "crypto/rand"
import "crypto/sha256"
func main() { key, _ := rsa.GenerateKey(rand.Reader, 512)
ct, _ := rsa.EncryptOAEP(sha256.New(), rand.Reader, &key.PublicKey, []byte("hi"), nil)
_, _ = rsa.DecryptOAEP(sha256.New(), rand.Reader, key, ct, nil) }
