// vybe-test: go/html_template_escape/html_escape_string_less_than
// origin: languages/go/tests/go/test_html_template_escape.rs

package main
import "fmt"
import "html"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(html.EscapeString("1<2")), "1&lt;2") }
