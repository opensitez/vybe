// vybe-test: go/url_query_extended/url_parse_request_uri_rejects_fragment
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

func main() { _, err := url.ParseRequestURI("/page#frag")
__check(fmt.Sprint(err != nil), "true") }
