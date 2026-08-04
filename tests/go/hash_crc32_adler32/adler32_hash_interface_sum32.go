// vybe-test: go/hash_crc32_adler32/adler32_hash_interface_sum32
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs

package main
import "fmt"
import "hash/adler32"
import "hash"
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

func main() { var h hash.Hash32 = adler32.New()
h.Write([]byte("go"))
__p(fmt.Sprint(h.Sum32())) 
__check("20906199")
}
