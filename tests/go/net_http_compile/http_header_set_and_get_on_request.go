// vybe-test: go/net_http_compile/http_header_set_and_get_on_request
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { req, _ := http.NewRequest(http.MethodGet, "https://example.com", nil)
req.Header.Set("Accept", "application/json")
_ = req.Header.Get("Accept") }
