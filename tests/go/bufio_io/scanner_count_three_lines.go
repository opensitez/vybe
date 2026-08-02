// vybe-test: go/bufio_io/scanner_count_three_lines
// origin: languages/go/tests/go/test_bufio_io.rs

package main
import "fmt"
import "bufio"
import "strings"
func main() { sc := bufio.NewScanner(strings.NewReader("one\ntwo\nthree\n"))
n := 0
for sc.Scan() { n++ }
fmt.Println(n) }
