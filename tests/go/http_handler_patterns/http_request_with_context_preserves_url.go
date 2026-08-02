// vybe-test: go/http_handler_patterns/http_request_with_context_preserves_url
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "context"
import "net/http"
func main() { req, _ := http.NewRequest(http.MethodGet, "https://ex.com/path", nil)
ctx := context.Background()
r2 := req.WithContext(ctx)
_ = r2.URL.Path }
