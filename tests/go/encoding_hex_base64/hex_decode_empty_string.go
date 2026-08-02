// vybe-test: go/encoding_hex_base64/hex_decode_empty_string
// origin: languages/go/tests/go/test_encoding_hex_base64.rs

package main
import "fmt"
import "encoding/hex"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, err := hex.DecodeString("")
__check(fmt.Sprint(len(b)), "0")
__check(fmt.Sprint(err == nil), "true") }
