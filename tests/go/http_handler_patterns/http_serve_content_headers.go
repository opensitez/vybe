// vybe-test: go/http_handler_patterns/http_serve_content_headers
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
import "strings"
import "time"
func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { http.ServeContent(w, r, "f.txt", time.Now(), strings.NewReader("data")) }) }
