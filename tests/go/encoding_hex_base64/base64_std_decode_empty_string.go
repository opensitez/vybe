// vybe-test: go/encoding_hex_base64/base64_std_decode_empty_string
// origin: languages/go/tests/go/test_encoding_hex_base64.rs

package main
import "fmt"
import "encoding/base64"
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

func main() { b, err := base64.StdEncoding.DecodeString("")
__p(fmt.Sprint(len(b)))
__p(fmt.Sprint(err == nil)) 
__check("0\ntrue")
}
