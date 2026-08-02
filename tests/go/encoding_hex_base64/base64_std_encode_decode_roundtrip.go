// vybe-test: go/encoding_hex_base64/base64_std_encode_decode_roundtrip
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

func main() { orig := []byte{0, 1, 255}
enc := base64.StdEncoding.EncodeToString(orig)
back, _ := base64.StdEncoding.DecodeString(enc)
__check(fmt.Sprint(len(back)), "3")
__check(fmt.Sprint(int(back[2])), "255") }
