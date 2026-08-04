// vybe-test: go/encoding_hex_base64/hex_dump_includes_offset_and_ascii
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

func main() { dump := string(hex.Dump([]byte("ab")))
__p(fmt.Sprint(len(dump) > 0))
__p(fmt.Sprint(dump[0:8])) 
__check("true\n00000000")
}
