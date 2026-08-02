// vybe-test: go/text_template_tabwriter/tabwriter_flush
// origin: languages/go/tests/go/test_text_template_tabwriter.rs
// vybe-test-mode: compile

package main
import "text/tabwriter"
import "bytes"
func main() { var b bytes.Buffer
w := tabwriter.NewWriter(&b, 0, 0, 1, ' ', 0)
w.Flush() }
