// vybe-test: go/io_pipe_copy_tee/io_tee_reader_nested_copy
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs
// vybe-test-mode: compile

package main
import "io"
import "bytes"
import "strings"
func main() { var side bytes.Buffer
tr := io.TeeReader(strings.NewReader("n"), &side)
var dst bytes.Buffer
_, _ = io.Copy(&dst, tr) }
