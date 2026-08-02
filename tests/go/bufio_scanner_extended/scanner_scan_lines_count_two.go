// vybe-test: go/bufio_scanner_extended/scanner_scan_lines_count_two
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs

package main
import "fmt"
import "bufio"
import "strings"
func main() { sc := bufio.NewScanner(strings.NewReader("x\ny\n"))
sc.Split(bufio.ScanLines)
n := 0
for sc.Scan() { n++ }
fmt.Println(n) }
