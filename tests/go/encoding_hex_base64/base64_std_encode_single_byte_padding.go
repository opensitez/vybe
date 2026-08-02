// vybe-test: go/encoding_hex_base64/base64_std_encode_single_byte_padding
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

func main() { __check(fmt.Sprint(base64.StdEncoding.EncodeToString([]byte("f"))), "Zg==") }
