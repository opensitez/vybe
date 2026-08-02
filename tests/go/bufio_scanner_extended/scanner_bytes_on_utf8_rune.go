// vybe-test: go/bufio_scanner_extended/scanner_bytes_on_utf8_rune
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs

package main
import "fmt"
import "bufio"
import "strings"
func main() { sc := bufio.NewScanner(strings.NewReader("日"))
sc.Split(bufio.ScanBytes)
n := 0
for sc.Scan() { n++ }
fmt.Println(n) }
