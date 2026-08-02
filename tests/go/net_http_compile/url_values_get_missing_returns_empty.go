// vybe-test: go/net_http_compile/url_values_get_missing_returns_empty
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { v := url.Values{}
_ = v.Get("missing") == "" }
