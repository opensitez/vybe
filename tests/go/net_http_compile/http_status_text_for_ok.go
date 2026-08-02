// vybe-test: go/net_http_compile/http_status_text_for_ok
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { _ = http.StatusText(http.StatusOK) }
