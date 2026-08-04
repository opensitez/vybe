// vybe-test: go/hash_crc32_adler32/crc32_update_incremental
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs

package main
import "fmt"
import "hash/crc32"
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

func main() { table := crc32.IEEETable
c := uint32(0)
c = crc32.Update(c, table, []byte("ab"))
c = crc32.Update(c, table, []byte("c"))
full := crc32.ChecksumIEEE([]byte("abc"))
__p(fmt.Sprint(c == full)) 
__check("true")
}
