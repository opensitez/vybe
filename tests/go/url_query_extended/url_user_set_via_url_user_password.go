// vybe-test: go/url_query_extended/url_user_set_via_url_user_password
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://host/")
u.User = url.UserPassword("admin", "pw")
_ = u.User.Username() }
