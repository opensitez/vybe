// vybe-test: go/bufio_scanner_extended/scanner_scan_bytes_count_three
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs

package main
import "fmt"
import "bufio"
import "strings"
func main() { sc := bufio.NewScanner(strings.NewReader("xyz"))
sc.Split(bufio.ScanBytes)
n := 0
for sc.Scan() { n++ }
fmt.Println(n) }
