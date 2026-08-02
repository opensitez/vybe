// vybe-test: go/net_http_compile/http_client_get_assign_response_fields
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { resp, err := http.DefaultClient.Get("https://api.test/items")
if err == nil { _ = resp.StatusCode
_ = resp.Header
_ = resp.Body } }
