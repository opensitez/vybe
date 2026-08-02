// vybe-test: go/url_query_extended/url_user_string_roundtrip
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

func main() { u, _ := url.Parse("https://user:pass@host/")
__check(fmt.Sprint(u.User.String()), "user:pass") }
