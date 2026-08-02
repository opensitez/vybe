// vybe-test: go/http_handler_patterns/http_response_writer_header_get_after_set
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { w.Header().Set("X-Trace", "1"); _ = w.Header().Get("X-Trace") }) }
