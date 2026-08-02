// vybe-test: go/bufio_scanner_extended/scanner_words_single_token_stream
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs

package main
import "fmt"
import "bufio"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { sc := bufio.NewScanner(strings.NewReader("solo"))
sc.Split(bufio.ScanWords)
sc.Scan()
__check(fmt.Sprint(sc.Text()), "solo")
__check(fmt.Sprint(sc.Scan()), "false") }
