// vybe-test: go/url_query_extended/url_query_add_empty_value
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { q := url.Values{}
q.Add("flag", "")
_ = q.Encode() }
