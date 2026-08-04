// vybe-test: go/encoding_binary/binary_little_endian_put_uint64_low_byte_first
// origin: languages/go/tests/go/test_encoding_binary.rs

package main
import "fmt"
import "encoding/binary"
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

func main() { buf := make([]byte, 8)
binary.LittleEndian.PutUint64(buf, 0x0102030405060708)
__p(fmt.Sprint(int(buf[0])))
__p(fmt.Sprint(int(buf[7]))) 
__check("8\n1")
}
