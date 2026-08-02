// vybe-test: go/http_handler_patterns/http_handler_func_satisfies_handler
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { var h http.Handler = http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {}) }
