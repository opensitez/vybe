// vybe-test: go/http_handler_patterns/http_handler_serve_http_method
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
type greeter struct{}
func (g greeter) ServeHTTP(w http.ResponseWriter, r *http.Request) {}
func main() { var h http.Handler = greeter{}
h.ServeHTTP(nil, nil) }
