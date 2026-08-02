// vybe-test: go/stdlib_crypto_extended/x509_system_cert_pool
// origin: languages/go/tests/go/test_stdlib_crypto_extended.rs
// vybe-test-mode: compile

package main
import "crypto/x509"
func main() { _ = x509.SystemCertPool() }
