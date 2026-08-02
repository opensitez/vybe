// vybe-test: go/url_query_extended/url_query_set_replaces_single_value
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

func main() { u, _ := url.Parse("https://host/")
q := u.Query()
q.Set("token", "abc")
u.RawQuery = q.Encode()
__check(fmt.Sprint(u.Query().Get("token")), "abc") }
