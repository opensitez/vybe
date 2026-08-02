// vybe-test: go/net_http_compile/http_status_ok_constant_compare
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { code := http.StatusOK
_ = code == 200 }
