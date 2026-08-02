// vybe-test: go/encoding_hex_base64/base64_std_encoded_len_formula
// origin: languages/go/tests/go/test_encoding_hex_base64.rs

package main
import "fmt"
import "encoding/base64"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(base64.StdEncoding.EncodedLen(1)), "4")
__check(fmt.Sprint(base64.StdEncoding.EncodedLen(3)), "4") }
