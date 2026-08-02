// vybe-test: go/net_http_compile/url_values_del_removes_key
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { v := url.Values{"k": []string{"1"}}
v.Del("k")
_ = len(v) }
