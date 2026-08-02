// vybe-test: go/encoding_hex_base64/hex_encode_empty_slice
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

func main() { src := []byte{}
dst := make([]byte, hex.EncodedLen(len(src)))
n := hex.Encode(dst, src)
__check(fmt.Sprint(n), "0") }
