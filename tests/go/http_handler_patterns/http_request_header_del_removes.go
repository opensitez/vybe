// vybe-test: go/http_handler_patterns/http_request_header_del_removes
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { req, _ := http.NewRequest(http.MethodGet, "https://ex.com", nil)
req.Header.Set("X-Tmp", "1")
req.Header.Del("X-Tmp")
_ = req.Header.Get("X-Tmp") }
