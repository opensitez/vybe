// vybe-test: go/cover_crypto_all/x509_create_certificate_request
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/x509"
import "crypto/rand"
import "crypto/rsa"
func main() { key, _ := rsa.GenerateKey(rand.Reader, 512)
_, _ = x509.CreateCertificateRequest(rand.Reader, []byte("csr"), key) }
