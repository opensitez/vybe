// vybe-test: go/net_http_compile/http_method_post_constant_in_request
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { _, _ = http.NewRequest(http.MethodPost, "https://example.com", nil) }
