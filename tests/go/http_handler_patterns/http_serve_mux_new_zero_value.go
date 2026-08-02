// vybe-test: go/http_handler_patterns/http_serve_mux_new_zero_value
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { var mux http.ServeMux
_ = mux }
