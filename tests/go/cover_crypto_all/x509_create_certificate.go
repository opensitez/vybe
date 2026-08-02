// vybe-test: go/cover_crypto_all/x509_create_certificate
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/x509"
import "crypto/rand"
import "crypto/rsa"
import "math/big"
func main() { key, _ := rsa.GenerateKey(rand.Reader, 512)
tmpl := x509.Certificate{SerialNumber: big.NewInt(1)}
_, _ = x509.CreateCertificate(rand.Reader, &tmpl, &tmpl, &key.PublicKey, key) }
