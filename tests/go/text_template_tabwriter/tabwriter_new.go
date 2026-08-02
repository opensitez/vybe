// vybe-test: go/text_template_tabwriter/tabwriter_new
// origin: languages/go/tests/go/test_text_template_tabwriter.rs
// vybe-test-mode: compile

package main
import "text/tabwriter"
func main() { w := tabwriter.NewWriter(nil, 0, 0, 1, ' ', 0)
_ = w }
