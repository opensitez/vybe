// vybe-test: go/net_http_compile/http_new_request_get_no_body
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { req, _ := http.NewRequest(http.MethodGet, "https://example.com", nil)
_ = req.URL }
