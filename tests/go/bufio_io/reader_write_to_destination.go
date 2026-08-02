// vybe-test: go/bufio_io/reader_write_to_destination
// origin: languages/go/tests/go/test_bufio_io.rs
// vybe-test-mode: compile

package main
import "bufio"
import "bytes"
import "strings"
func main() { r := bufio.NewReader(strings.NewReader("go"))
var dst bytes.Buffer
_, _ = r.WriteTo(&dst) }
