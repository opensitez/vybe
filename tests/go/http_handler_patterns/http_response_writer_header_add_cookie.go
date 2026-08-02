// vybe-test: go/http_handler_patterns/http_response_writer_header_add_cookie
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { w.Header().Add("Set-Cookie", "sid=abc") }) }
