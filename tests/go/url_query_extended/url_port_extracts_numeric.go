// vybe-test: go/url_query_extended/url_port_extracts_numeric
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://api.test:9090/v1")
_ = u.Port() }
