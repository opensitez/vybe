// vybe-test: go/bufio_scanner_extended/scanner_custom_split_fixed_width_two
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { sc := bufio.NewScanner(strings.NewReader("abcdef"))
sc.Split(func(data []byte, atEOF bool) (advance int, token []byte, err error) { if len(data) >= 2 { return 2, data[:2], nil }; return 0, data, nil })
_ = sc.Scan() }
