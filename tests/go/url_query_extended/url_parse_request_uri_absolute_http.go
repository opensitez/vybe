// vybe-test: go/url_query_extended/url_parse_request_uri_absolute_http
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

func main() { u, _ := url.ParseRequestURI("http://example.com/health")
__check(fmt.Sprint(u.Scheme), "http")
__check(fmt.Sprint(u.Host), "example.com") }
