// vybe-test: go/html_template_escape/html_template_execute_escaped_pipeline
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

func main() { t := template.Must(template.New("p").Parse(`{{. | html}}`))
var b bytes.Buffer
t.Execute(&b, "<a>")
__check(fmt.Sprint(b.String()), "&lt;a&gt;") }
