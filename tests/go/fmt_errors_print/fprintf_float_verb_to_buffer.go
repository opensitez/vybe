// vybe-test: go/fmt_errors_print/fprintf_float_verb_to_buffer
// origin: languages/go/tests/go/test_fmt_errors_print.rs
// vybe-test-mode: compile

package main
import "fmt"
import "bytes"
func main() { var buf bytes.Buffer
_, _ = fmt.Fprintf(&buf, "%.1f", 1.25) }
