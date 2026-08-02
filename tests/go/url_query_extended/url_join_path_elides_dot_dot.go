// vybe-test: go/url_query_extended/url_join_path_elides_dot_dot
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

func main() { __check(fmt.Sprint(url.JoinPath("/a/b", "../c")), "/a/c") }
