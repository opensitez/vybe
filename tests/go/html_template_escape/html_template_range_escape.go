// vybe-test: go/html_template_escape/html_template_range_escape
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

func main() { t := template.Must(template.New("p").Parse(`{{range .}}{{.}}{{end}}`))
var b bytes.Buffer
t.Execute(&b, []string{"<a>", "b"})
__check(fmt.Sprint(b.String()), "&lt;a&gt;b") }
