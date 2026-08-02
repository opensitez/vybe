// vybe-test: go/url_query_extended/url_parse_then_mutate_scheme
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("http://h/")
u.Scheme = "https"
_ = u.Scheme }
