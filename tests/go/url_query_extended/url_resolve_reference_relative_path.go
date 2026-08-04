// vybe-test: go/url_query_extended/url_resolve_reference_relative_path
// origin: languages/go/tests/go/test_url_query_extended.rs

package main
import "fmt"
import "net/url"
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { base, _ := url.Parse("https://example.com/a/b")
ref, _ := url.Parse("c")
__p(fmt.Sprint(base.ResolveReference(ref).String())) 
__check("https://example.com/a/c")
}
