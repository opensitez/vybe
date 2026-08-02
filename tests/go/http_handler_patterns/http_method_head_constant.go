// vybe-test: go/http_handler_patterns/http_method_head_constant
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { _, _ = http.NewRequest(http.MethodHead, "https://ex.com/item/1", nil) }
