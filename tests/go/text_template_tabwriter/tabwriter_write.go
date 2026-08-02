// vybe-test: go/text_template_tabwriter/tabwriter_write
// origin: languages/go/tests/go/test_text_template_tabwriter.rs
// vybe-test-mode: compile

package main
import "text/tabwriter"
import "os"
func main() { w := tabwriter.NewWriter(os.Stdout, 0, 0, 1, ' ', 0)
_, _ = w.Write([]byte("a\tb\n")) }
