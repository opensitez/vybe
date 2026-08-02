// vybe-test: go/strings_bytes_compare/bytes_equal_same_content_diff_cap
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

func main() { a := append([]byte{}, 'x')
b := []byte{'x'}
__check(fmt.Sprint(bytes.Equal(a, b)), "true") }
