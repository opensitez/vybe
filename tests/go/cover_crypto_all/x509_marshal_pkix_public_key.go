// vybe-test: go/cover_crypto_all/x509_marshal_pkix_public_key
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/x509"
import "crypto/rsa"
import "crypto/rand"
func main() { key, _ := rsa.GenerateKey(rand.Reader, 512)
_, _ = x509.MarshalPKIXPublicKey(&key.PublicKey) }
