// vybe-test: go/net_http_compile/url_path_unescape_encoded_segments
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { _ = url.PathUnescape("a%2Fb%20c") }
