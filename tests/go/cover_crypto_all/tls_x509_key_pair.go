// vybe-test: go/cover_crypto_all/tls_x509_key_pair
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/tls"
func main() { _, _ = tls.X509KeyPair([]byte("cert"), []byte("key")) }
