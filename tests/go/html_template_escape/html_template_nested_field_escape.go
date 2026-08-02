// vybe-test: go/html_template_escape/html_template_nested_field_escape
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

func main() { type Inner struct { V string }
type Outer struct { Inner Inner }
t := template.Must(template.New("p").Parse("{{.Inner.V}}"))
var b bytes.Buffer
t.Execute(&b, Outer{Inner: Inner{V: "<z>"}})
__check(fmt.Sprint(b.String()), "&lt;z&gt;") }
