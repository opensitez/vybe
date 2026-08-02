// vybe-test: go/http_handler_patterns/http_request_basic_auth_set
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { req, _ := http.NewRequest(http.MethodGet, "https://ex.com", nil)
req.SetBasicAuth("user", "pass")
_ = req.Header.Get("Authorization") }
