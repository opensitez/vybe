// vybe-test: go/http_handler_patterns/http_request_user_agent_set
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { req, _ := http.NewRequest(http.MethodGet, "https://ex.com", nil)
req.Header.Set("User-Agent", "vybe-test/1.0")
_ = req.Header.Get("User-Agent") }
