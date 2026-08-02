// vybe-test: go/cover_crypto_all/tls_dial
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/tls"
func main() { _, _ = tls.Dial("tcp", "example.com:443", nil) }
