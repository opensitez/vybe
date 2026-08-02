// vybe-test: go/cover_crypto_all/tls_new_listener
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/tls"
import "net"
func main() { ln, _ := net.Listen("tcp", ":8443")
_ = tls.NewListener(ln, nil) }
