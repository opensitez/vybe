// vybe-test: go/url_query_extended/url_resolve_reference_relative_path
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

func main() { base, _ := url.Parse("https://example.com/a/b")
ref, _ := url.Parse("c")
__check(fmt.Sprint(base.ResolveReference(ref).String()), "https://example.com/a/c") }
