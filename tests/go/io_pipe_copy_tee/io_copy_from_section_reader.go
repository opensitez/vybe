// vybe-test: go/io_pipe_copy_tee/io_copy_from_section_reader
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs
// vybe-test-mode: compile

package main
import "io"
import "bytes"
import "strings"
func main() { sr := strings.NewReader("section")
var dst bytes.Buffer
_, _ = io.Copy(&dst, io.LimitReader(sr, 4)) }
