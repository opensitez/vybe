// vybe-test: go/function_literals_closures/closure_handler_func_serve_http_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { _, _ = w.Write([]byte("ok")) }).ServeHTTP(nil, nil) }
