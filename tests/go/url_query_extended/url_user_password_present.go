// vybe-test: go/url_query_extended/url_user_password_present
// origin: languages/go/tests/go/test_url_query_extended.rs

package main
import "fmt"
import "net/url"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { u, _ := url.Parse("https://bob:secret@host/")
pw, ok := u.User.Password()
__check(fmt.Sprint(pw), "secret")
__check(fmt.Sprint(ok), "true") }
