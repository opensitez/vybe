// vybe-test: go/net_http_compile/http_new_request_post_with_reader
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/http"
import "strings"
func main() { body := strings.NewReader("payload")
_, _ = http.NewRequest(http.MethodPost, "https://example.com/submit", body) }
