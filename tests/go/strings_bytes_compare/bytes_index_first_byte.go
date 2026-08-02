// vybe-test: go/strings_bytes_compare/bytes_index_first_byte
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

func main() { __check(fmt.Sprint(bytes.Index([]byte("abc"), []byte("a"))), "0") }
