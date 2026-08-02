// vybe-test: go/html_template_escape/html_template_execute_struct_field
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

func main() { type Page struct { Title string }
t := template.Must(template.New("p").Parse("{{.Title}}"))
var b bytes.Buffer
t.Execute(&b, Page{Title: "<x>"})
__check(fmt.Sprint(b.String()), "&lt;x&gt;") }
