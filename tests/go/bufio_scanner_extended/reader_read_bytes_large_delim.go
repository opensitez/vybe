// vybe-test: go/bufio_scanner_extended/reader_read_bytes_large_delim
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { r := bufio.NewReader(strings.NewReader("startEND"))
_, _ = r.ReadBytes('E') }
