// vybe-test: go/encoding_hex_base64/hex_encode_decode_roundtrip
// origin: languages/go/tests/go/test_encoding_hex_base64.rs

package main
import "fmt"
import "encoding/hex"
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { orig := []byte{1, 2, 3, 250}
enc := hex.EncodeToString(orig)
back, _ := hex.DecodeString(enc)
__p(fmt.Sprint(len(back)))
__p(fmt.Sprint(int(back[3]))) 
__check("4\n250")
}
