// vybe-test: go/bytes_buffer_extended/buffer_readfrom_empty_reader
// origin: languages/go/tests/go/test_bytes_buffer_extended.rs
// vybe-test-mode: compile

package main
import "bytes"
import "strings"
func main() { var b bytes.Buffer
_, _ = b.ReadFrom(strings.NewReader("")) }
