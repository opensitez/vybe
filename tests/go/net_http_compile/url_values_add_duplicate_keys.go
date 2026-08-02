// vybe-test: go/net_http_compile/url_values_add_duplicate_keys
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { v := url.Values{}
v.Add("tag", "a")
v.Add("tag", "b")
_ = v.Encode() }
