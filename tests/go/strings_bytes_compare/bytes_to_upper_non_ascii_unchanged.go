// vybe-test: go/strings_bytes_compare/bytes_to_upper_non_ascii_unchanged
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

func main() { b := []byte("日")
u := bytes.ToUpper(b)
__check(fmt.Sprint(bytes.Equal(b, u)), "true") }
