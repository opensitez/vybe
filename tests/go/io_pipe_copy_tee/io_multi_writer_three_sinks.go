// vybe-test: go/io_pipe_copy_tee/io_multi_writer_three_sinks
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs
// vybe-test-mode: compile

package main
import "io"
import "bytes"
func main() { var a bytes.Buffer
var b bytes.Buffer
var c bytes.Buffer
mw := io.MultiWriter(&a, &b, &c)
_, _ = mw.Write([]byte("z")) }
