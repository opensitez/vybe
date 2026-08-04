// vybe-test: go/encoding_binary/binary_read_uint16_from_bytes_reader
// origin: languages/go/tests/go/test_encoding_binary.rs

package main
import "bytes"
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

func main() { r := bytes.NewReader([]byte{0x01, 0x02})
var v uint16
_ = binary.Read(r, binary.BigEndian, &v)
__p(fmt.Sprint(v)) 
__check("258")
}
