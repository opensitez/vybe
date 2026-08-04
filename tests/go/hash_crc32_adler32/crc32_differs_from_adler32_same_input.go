// vybe-test: go/hash_crc32_adler32/crc32_differs_from_adler32_same_input
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs

package main
import "fmt"
import "hash/crc32"
import "hash/adler32"
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

func main() { data := []byte("test")
c := crc32.ChecksumIEEE(data)
a := adler32.Checksum(data)
__p(fmt.Sprint(c != uint32(a))) 
__check("true")
}
