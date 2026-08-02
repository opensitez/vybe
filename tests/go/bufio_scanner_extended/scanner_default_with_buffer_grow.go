// vybe-test: go/bufio_scanner_extended/scanner_default_with_buffer_grow
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { sc := bufio.NewScanner(strings.NewReader("longline"))
sc.Buffer(make([]byte, 4), 128)
for sc.Scan() { _ = sc.Text() } }
