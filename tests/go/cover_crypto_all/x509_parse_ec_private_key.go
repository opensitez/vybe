// vybe-test: go/cover_crypto_all/x509_parse_ec_private_key
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/x509"
func main() { _, _ = x509.ParseECPrivateKey(nil) }
