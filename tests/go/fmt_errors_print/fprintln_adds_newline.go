// vybe-test: go/fmt_errors_print/fprintln_adds_newline
// origin: languages/go/tests/go/test_fmt_errors_print.rs

package main
import "fmt"
import "bytes"
func main() { var buf bytes.Buffer
fmt.Fprintln(&buf, "go")
fmt.Println(buf.String()) }
