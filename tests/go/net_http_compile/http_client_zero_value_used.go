// vybe-test: go/net_http_compile/http_client_zero_value_used
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { var c http.Client
_, _ = c.Get("https://example.com") }
