// vybe-test: go/function_literals_closures/closure_as_http_handler_func_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { var h http.Handler = http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { w.WriteHeader(http.StatusOK) }) }
