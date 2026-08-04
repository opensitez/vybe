// vybe-test: go/url_query_extended/url_query_get_from_parsed_url
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

func main() { u, _ := url.Parse("https://host/search?q=vybe&lang=go")
__p(fmt.Sprint(u.Query().Get("q")))
__p(fmt.Sprint(u.Query().Get("lang"))) 
__check("vybe\ngo")
}
