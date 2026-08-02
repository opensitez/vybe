// vybe-test: go/url_query_extended/url_is_abs_on_relative_path
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("/rel")
_ = u.IsAbs() }
