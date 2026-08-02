// vybe-test: go/encoding_hex_base64/hex_encode_decode_roundtrip
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

func main() { orig := []byte{1, 2, 3, 250}
enc := hex.EncodeToString(orig)
back, _ := hex.DecodeString(enc)
__check(fmt.Sprint(len(back)), "4")
__check(fmt.Sprint(int(back[3])), "250") }
