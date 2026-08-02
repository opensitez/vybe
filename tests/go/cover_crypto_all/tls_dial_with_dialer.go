// vybe-test: go/cover_crypto_all/tls_dial_with_dialer
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/tls"
import "net"
func main() { d := net.Dialer{}
_, _ = tls.DialWithDialer(&d, "tcp", "example.com:443", nil) }
