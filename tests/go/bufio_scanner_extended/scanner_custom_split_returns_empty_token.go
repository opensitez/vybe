// vybe-test: go/bufio_scanner_extended/scanner_custom_split_returns_empty_token
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { sc := bufio.NewScanner(strings.NewReader(",a"))
sc.Split(func(data []byte, atEOF bool) (advance int, token []byte, err error) { if len(data) > 0 && data[0] == ',' { return 1, []byte{}, nil }; return 0, nil, nil })
_ = sc.Scan() }
