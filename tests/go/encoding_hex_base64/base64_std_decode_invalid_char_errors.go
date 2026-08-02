// vybe-test: go/encoding_hex_base64/base64_std_decode_invalid_char_errors
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

func main() { _, err := base64.StdEncoding.DecodeString("!!!")
__check(fmt.Sprint(err != nil), "true") }
