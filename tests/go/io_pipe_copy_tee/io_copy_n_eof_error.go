// vybe-test: go/io_pipe_copy_tee/io_copy_n_eof_error
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs
// vybe-test-mode: compile

package main
import "io"
import "bytes"
import "strings"
func main() { var dst bytes.Buffer
_, _ = io.CopyN(&dst, strings.NewReader("a"), 3) }
