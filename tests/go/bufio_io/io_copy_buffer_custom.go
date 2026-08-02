// vybe-test: go/bufio_io/io_copy_buffer_custom
// origin: languages/go/tests/go/test_bufio_io.rs
// vybe-test-mode: compile

package main
import "io"
import "bytes"
import "strings"
func main() { var dst bytes.Buffer
buf := make([]byte, 8)
_, _ = io.CopyBuffer(&dst, strings.NewReader("buf"), buf) }
