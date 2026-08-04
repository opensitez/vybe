// vybe-test: go/fmt_errors_print/fprintf_formatted_to_buffer
// origin: languages/go/tests/go/test_fmt_errors_print.rs

package main
import "fmt"
import "bytes"
func main() { var buf bytes.Buffer
fmt.Fprintf(&buf, "id=%d", 7)
fmt.Println(buf.String()) }
