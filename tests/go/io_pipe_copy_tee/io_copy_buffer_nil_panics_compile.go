// vybe-test: go/io_pipe_copy_tee/io_copy_buffer_nil_panics_compile
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs
// vybe-test-mode: compile

package main
import "io"
import "bytes"
import "strings"
func main() { var dst bytes.Buffer
_, _ = io.CopyBuffer(&dst, strings.NewReader("x"), nil) }
