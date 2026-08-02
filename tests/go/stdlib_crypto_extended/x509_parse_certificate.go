// vybe-test: go/stdlib_crypto_extended/x509_parse_certificate
// origin: languages/go/tests/go/test_stdlib_crypto_extended.rs
// vybe-test-mode: compile

package main
import "crypto/x509"
func main() { _, _ = x509.ParseCertificate(nil) }
