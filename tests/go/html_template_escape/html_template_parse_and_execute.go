// vybe-test: go/html_template_escape/html_template_parse_and_execute
// origin: languages/go/tests/go/test_html_template_escape.rs

package main
import "fmt"
import "html/template"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { t, _ := template.New("x").Parse("{{.N}}")
var b bytes.Buffer
t.Execute(&b, struct{ N int }{7})
__check(fmt.Sprint(b.String()), "7") }
