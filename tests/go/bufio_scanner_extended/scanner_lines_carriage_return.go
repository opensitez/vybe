// vybe-test: go/bufio_scanner_extended/scanner_lines_carriage_return
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { sc := bufio.NewScanner(strings.NewReader("a\r\nb"))
sc.Split(bufio.ScanLines)
for sc.Scan() { _ = sc.Text() } }
