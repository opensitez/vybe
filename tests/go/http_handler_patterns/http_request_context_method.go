// vybe-test: go/http_handler_patterns/http_request_context_method
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "context"
import "net/http"
func main() { req, _ := http.NewRequest(http.MethodGet, "https://ex.com", nil)
req = req.WithContext(context.Background())
_ = req.Context() }
