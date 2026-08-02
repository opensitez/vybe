// vybe-test: go/bufio_scanner_extended/scanner_split_reassign_after_scan
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { sc := bufio.NewScanner(strings.NewReader("a b"))
sc.Split(bufio.ScanWords)
_ = sc.Scan()
sc.Split(bufio.ScanLines)
_ = sc.Scan() }
