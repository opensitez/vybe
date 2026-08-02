// vybe-test: go/bufio_io/writer_new_writer_size
// origin: languages/go/tests/go/test_bufio_io.rs
// vybe-test-mode: compile

package main
import "bufio"
import "bytes"
func main() { var buf bytes.Buffer
w := bufio.NewWriterSize(&buf, 16)
w.WriteString("size") }
