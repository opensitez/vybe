// vybe-test: go/http_handler_patterns/http_request_with_context_returns_new_request
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "context"
import "net/http"
func main() { req, _ := http.NewRequest(http.MethodGet, "https://ex.com", nil)
ctx := context.WithValue(context.Background(), "k", 1)
_ = req.WithContext(ctx) }
