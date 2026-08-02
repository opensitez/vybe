// vybe-test: go/net_http_compile/http_status_internal_server_error_constant
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { _ = http.StatusInternalServerError == 500 }
