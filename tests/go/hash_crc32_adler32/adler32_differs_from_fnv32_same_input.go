// vybe-test: go/hash_crc32_adler32/adler32_differs_from_fnv32_same_input
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs

package main
import "fmt"
import "hash/adler32"
import "hash/fnv"
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
a := adler32.Checksum(data)
h := fnv.New32a()
h.Write(data)
__p(fmt.Sprint(uint32(a) != h.Sum32())) 
__check("true")
}
