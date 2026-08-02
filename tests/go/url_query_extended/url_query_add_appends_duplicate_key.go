// vybe-test: go/url_query_extended/url_query_add_appends_duplicate_key
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
q.Add("id", "1")
q.Add("id", "2")
__check(fmt.Sprint(len(q["id"])), "2") }
