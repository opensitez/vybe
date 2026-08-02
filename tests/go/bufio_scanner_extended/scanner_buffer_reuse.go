// vybe-test: go/bufio_scanner_extended/scanner_buffer_reuse
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { sc := bufio.NewScanner(strings.NewReader("tok"))
_ = sc.Buffer(make([]byte, 0, 64), 1024)
_ = sc.Scan() }
