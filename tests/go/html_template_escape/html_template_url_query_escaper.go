// vybe-test: go/html_template_escape/html_template_url_query_escaper
// origin: languages/go/tests/go/test_html_template_escape.rs

package main
import "fmt"
import "html/template"
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

func main() { s := template.URLQueryEscaper([]byte("q=go lang"))
__p(fmt.Sprint(len(s) > 0)) 
__check("true")
}
