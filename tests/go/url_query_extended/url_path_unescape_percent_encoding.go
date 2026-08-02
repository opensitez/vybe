// vybe-test: go/url_query_extended/url_path_unescape_percent_encoding
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

func main() { s, _ := url.PathUnescape("a%2Fb%20c")
__check(fmt.Sprint(s), "a/b c") }
