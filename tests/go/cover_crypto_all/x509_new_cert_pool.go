// vybe-test: go/cover_crypto_all/x509_new_cert_pool
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/x509"
func main() { p := x509.NewCertPool()
p.AddCert(nil) }
