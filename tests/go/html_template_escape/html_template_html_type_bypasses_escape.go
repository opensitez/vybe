// vybe-test: go/html_template_escape/html_template_html_type_bypasses_escape
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

func main() { t := template.Must(template.New("p").Parse(`{{.}}`))
var b bytes.Buffer
t.Execute(&b, template.HTML("<em>hi</em>"))
__check(fmt.Sprint(b.String()), "<em>hi</em>") }
