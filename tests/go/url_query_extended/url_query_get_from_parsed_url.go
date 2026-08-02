// vybe-test: go/url_query_extended/url_query_get_from_parsed_url
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

func main() { u, _ := url.Parse("https://host/search?q=vybe&lang=go")
__check(fmt.Sprint(u.Query().Get("q")), "vybe")
__check(fmt.Sprint(u.Query().Get("lang")), "go") }
