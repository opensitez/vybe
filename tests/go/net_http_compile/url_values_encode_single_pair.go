// vybe-test: go/net_http_compile/url_values_encode_single_pair
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { v := url.Values{}
v.Set("q", "golang")
_ = v.Encode() }
