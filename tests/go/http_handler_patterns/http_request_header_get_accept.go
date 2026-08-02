// vybe-test: go/http_handler_patterns/http_request_header_get_accept
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { req, _ := http.NewRequest(http.MethodGet, "https://ex.com", nil)
req.Header.Set("Accept", "application/json")
_ = req.Header.Get("Accept") }
