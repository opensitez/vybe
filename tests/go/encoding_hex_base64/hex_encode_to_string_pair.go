// vybe-test: go/encoding_hex_base64/hex_encode_to_string_pair
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

func main() { __check(fmt.Sprint(hex.EncodeToString([]byte{0x0a, 0x0b})), "0a0b") }
