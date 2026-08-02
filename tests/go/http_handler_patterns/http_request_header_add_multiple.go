// vybe-test: go/http_handler_patterns/http_request_header_add_multiple
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { req, _ := http.NewRequest(http.MethodGet, "https://ex.com", nil)
req.Header.Add("X-Custom", "a")
req.Header.Add("X-Custom", "b")
_ = req.Header["X-Custom"] }
