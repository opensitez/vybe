// vybe-test: go/url_query_extended/url_user_string_without_password
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://solo@host/")
_ = u.User.String() }
