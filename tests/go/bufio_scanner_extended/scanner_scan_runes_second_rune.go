// vybe-test: go/bufio_scanner_extended/scanner_scan_runes_second_rune
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { sc := bufio.NewScanner(strings.NewReader("ab"))
sc.Split(bufio.ScanRunes)
_ = sc.Scan()
_ = sc.Scan() }
