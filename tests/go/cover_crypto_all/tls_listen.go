// vybe-test: go/cover_crypto_all/tls_listen
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/tls"
func main() { _, _ = tls.Listen("tcp", ":8443", nil) }
