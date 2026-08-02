// vybe-test: go/url_query_extended/url_parse_request_uri_path_and_query
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

func main() { u, _ := url.ParseRequestURI("/api/v2/items?page=3")
__check(fmt.Sprint(u.Path), "/api/v2/items")
__check(fmt.Sprint(u.Query().Get("page")), "3") }
