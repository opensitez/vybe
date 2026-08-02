// vybe-test: go/url_query_extended/url_resolve_reference_absolute_override
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

func main() { base, _ := url.Parse("https://example.com/old")
ref, _ := url.Parse("https://other/new")
__check(fmt.Sprint(base.ResolveReference(ref).String()), "https://other/new") }
