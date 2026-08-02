// vybe-test: go/function_literals_closures/closure_handler_func_with_capture_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { prefix := "x"
http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { _ = prefix + r.Method }) }
