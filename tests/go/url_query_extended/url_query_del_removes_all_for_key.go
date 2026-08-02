// vybe-test: go/url_query_extended/url_query_del_removes_all_for_key
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

func main() { u, _ := url.Parse("https://host/?x=1&x=2")
q := u.Query()
q.Del("x")
__check(fmt.Sprint(len(q)), "0") }
