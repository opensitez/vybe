// vybe-test: go/url_query_extended/url_query_set_on_existing_overwrites
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://h/?x=old")
q := u.Query()
q.Set("x", "new")
u.RawQuery = q.Encode()
_ = u.Query().Get("x") }
