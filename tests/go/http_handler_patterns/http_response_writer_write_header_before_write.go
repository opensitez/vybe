// vybe-test: go/http_handler_patterns/http_response_writer_write_header_before_write
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { w.WriteHeader(http.StatusCreated); _, _ = w.Write([]byte("created")) }) }
