// vybe-test: go/http_handler_patterns/http_response_writer_flusher_interface
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { if f, ok := w.(http.Flusher); ok { f.Flush() } }) }
