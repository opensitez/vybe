// vybe-test: go/url_query_extended/url_query_encode_special_characters
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { q := url.Values{}
q.Set("q", "a&b=c")
_ = q.Encode() }
