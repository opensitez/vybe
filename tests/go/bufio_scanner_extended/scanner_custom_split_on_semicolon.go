// vybe-test: go/bufio_scanner_extended/scanner_custom_split_on_semicolon
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { sc := bufio.NewScanner(strings.NewReader("a;b"))
sc.Split(func(data []byte, atEOF bool) (advance int, token []byte, err error) { for i, b := range data { if b == ';' { return i + 1, data[:i], nil } }; if atEOF && len(data) > 0 { return len(data), data, nil }; return 0, nil, nil })
_ = sc.Scan() }
