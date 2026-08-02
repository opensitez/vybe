// vybe-test: go/cover_crypto_all/tls_new_conn
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/tls"
func main() { _ = tls.NewConn(nil, nil) }
