// vybe-test: go/html_template_escape/html_template_css_escape_string
// origin: languages/go/tests/go/test_html_template_escape.rs

package main
import "fmt"
import "html/template"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := template.CSSEscapeString("<style>")
__check(fmt.Sprint(len(s) > 0), "true") }
