// vybe-test: go/stdlib_crypto_extended/tls_dial_compile
// origin: languages/go/tests/go/test_stdlib_crypto_extended.rs
// vybe-test-mode: compile

package main
import "crypto/tls"
func main() { _, _ = tls.Dial("tcp", "example.com:443", nil) }
