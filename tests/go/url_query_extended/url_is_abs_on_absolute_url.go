// vybe-test: go/url_query_extended/url_is_abs_on_absolute_url
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://a/b")
_ = u.IsAbs() }
