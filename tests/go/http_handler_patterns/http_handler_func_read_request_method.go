// vybe-test: go/http_handler_patterns/http_handler_func_read_request_method
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { _ = r.Method }) }
