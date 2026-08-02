// vybe-test: go/url_query_extended/url_query_del_one_of_many_keys
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://h/?a=1&b=2")
q := u.Query()
q.Del("a")
_ = q.Get("b") }
