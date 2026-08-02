// vybe-test: go/stdlib_mime_runtime/html_template_escape
// origin: languages/go/tests/go/test_stdlib_mime_runtime.rs
// vybe-test-mode: compile

package main
import "html/template"
func main() { _ = html/template.HTMLEscapeString("<b>") }
