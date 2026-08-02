// vybe-test: go/bufio_scanner_extended/scanner_custom_split_comma_with_remainder
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { sc := bufio.NewScanner(strings.NewReader("x,y,z"))
sc.Split(func(data []byte, atEOF bool) (advance int, token []byte, err error) { for i, b := range data { if b == ',' { return i + 1, data[:i], nil } }; return 0, nil, nil })
for sc.Scan() { _ = sc.Text() } }
