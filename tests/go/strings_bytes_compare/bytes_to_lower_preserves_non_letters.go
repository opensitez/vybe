// vybe-test: go/strings_bytes_compare/bytes_to_lower_preserves_non_letters
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

func main() { __check(fmt.Sprint(string(bytes.ToLower([]byte("A1_B")))), "a1_b") }
