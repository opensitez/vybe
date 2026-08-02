// vybe-test: go/url_query_extended/url_raw_query_roundtrip_via_encode
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

func main() { u, _ := url.Parse("https://h/?k=v")
q := u.Query()
u.RawQuery = q.Encode()
__check(fmt.Sprint(u.RawQuery), "k=v") }
