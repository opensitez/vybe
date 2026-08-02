// vybe-test: go/http_handler_patterns/http_close_notifier_deprecated_compile
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { _, _ = w.(http.CloseNotifier) }) }
