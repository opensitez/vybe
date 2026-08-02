// vybe-test: go/http_handler_patterns/http_handler_chain_middleware_pattern
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { wrap := func(next http.Handler) http.Handler { return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { next.ServeHTTP(w, r) }) }
_ = wrap(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {})) }
