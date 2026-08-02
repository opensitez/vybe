// vybe-test: go/bufio_io/writer_write_rune
// origin: languages/go/tests/go/test_bufio_io.rs
// vybe-test-mode: compile

package main
import "bufio"
import "bytes"
func main() { var buf bytes.Buffer
w := bufio.NewWriter(&buf)
_, _ = w.WriteRune('日')
w.Flush() }
