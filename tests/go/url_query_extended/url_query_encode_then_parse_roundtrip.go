// vybe-test: go/url_query_extended/url_query_encode_then_parse_roundtrip
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { q := url.Values{}
q.Set("z", "9")
u, _ := url.Parse("https://h/?" + q.Encode())
_ = u.Query().Get("z") }
