// vybe-test: go/bufio_io/io_multi_writer_duplicates
// origin: languages/go/tests/go/test_bufio_io.rs
// vybe-test-mode: compile

package main
import "io"
import "bytes"
func main() { var a bytes.Buffer
var b bytes.Buffer
mw := io.MultiWriter(&a, &b)
_, _ = mw.Write([]byte("x")) }
