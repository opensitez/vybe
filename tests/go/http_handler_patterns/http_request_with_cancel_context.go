// vybe-test: go/http_handler_patterns/http_request_with_cancel_context
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "context"
import "net/http"
func main() { ctx, cancel := context.WithCancel(context.Background())
defer cancel()
req, _ := http.NewRequestWithContext(ctx, http.MethodGet, "https://ex.com", nil)
_ = req.Context().Err() }
