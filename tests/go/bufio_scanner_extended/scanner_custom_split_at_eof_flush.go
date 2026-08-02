// vybe-test: go/bufio_scanner_extended/scanner_custom_split_at_eof_flush
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { sc := bufio.NewScanner(strings.NewReader("end"))
sc.Split(func(data []byte, atEOF bool) (advance int, token []byte, err error) { if atEOF { return len(data), data, nil }; return 0, nil, nil })
_ = sc.Scan() }
