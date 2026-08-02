// vybe-test: go/cover_crypto_all/tls_load_x509_key_pair
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/tls"
func main() { _, _ = tls.LoadX509KeyPair("cert.pem", "key.pem") }
