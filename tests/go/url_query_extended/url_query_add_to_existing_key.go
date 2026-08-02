// vybe-test: go/url_query_extended/url_query_add_to_existing_key
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://h/?k=1")
q := u.Query()
q.Add("k", "2")
_ = len(q["k"]) }
