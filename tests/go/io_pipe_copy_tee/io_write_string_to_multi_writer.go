// vybe-test: go/io_pipe_copy_tee/io_write_string_to_multi_writer
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs
// vybe-test-mode: compile

package main
import "io"
import "bytes"
func main() { var a bytes.Buffer
var b bytes.Buffer
mw := io.MultiWriter(&a, &b)
_, _ = io.WriteString(mw, "mw") }
