// vybe-test: go/url_query_extended/url_user_password_false_when_absent
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://guest@host/")
_, ok := u.User.Password()
_ = ok }
