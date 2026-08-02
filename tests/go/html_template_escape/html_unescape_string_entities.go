// vybe-test: go/html_template_escape/html_unescape_string_entities
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

func main() { __check(fmt.Sprint(html.UnescapeString("&lt;b&gt;")), "<b>") }
