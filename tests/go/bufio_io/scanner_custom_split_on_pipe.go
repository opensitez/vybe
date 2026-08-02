// vybe-test: go/bufio_io/scanner_custom_split_on_pipe
// origin: languages/go/tests/go/test_bufio_io.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { sc := bufio.NewScanner(strings.NewReader("a|b"))
sc.Split(func(data []byte, atEOF bool) (advance int, token []byte, err error) { for i, b := range data { if b == '|' { return i + 1, data[:i], nil } }; return 0, nil, nil })
_ = sc.Scan() }
