// vybe-test: go/url_query_extended/url_query_encode_sorted_keys
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

func main() { q := url.Values{}
q.Set("b", "2")
q.Set("a", "1")
__check(fmt.Sprint(q.Encode()), "a=1&b=2") }
