// vybe-test: go/net_http_compile/http_get_package_level_function
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { _, _ = http.Get("https://example.com/health") }
