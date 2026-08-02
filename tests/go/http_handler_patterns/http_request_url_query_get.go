// vybe-test: go/http_handler_patterns/http_request_url_query_get
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { req, _ := http.NewRequest(http.MethodGet, "https://ex.com/?q=go", nil)
_ = req.URL.Query().Get("q") }
