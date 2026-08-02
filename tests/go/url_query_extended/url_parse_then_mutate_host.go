// vybe-test: go/url_query_extended/url_parse_then_mutate_host
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://old.com/x")
u.Host = "new.com"
_ = u.Host }
