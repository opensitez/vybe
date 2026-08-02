// vybe-test: go/net_http_compile/url_query_unescape_percent_plus
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { _ = url.QueryUnescape("hello+world%21") }
