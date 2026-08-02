// vybe-test: go/strings_bytes_compare/bytes_compare_nil_vs_empty
// origin: languages/go/tests/go/test_strings_bytes_compare.rs

package main
import "fmt"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(bytes.Compare(nil, []byte{})), "0") }
