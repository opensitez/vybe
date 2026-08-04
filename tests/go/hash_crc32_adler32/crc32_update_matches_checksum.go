// vybe-test: go/hash_crc32_adler32/crc32_update_matches_checksum
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
direct := crc32.ChecksumIEEE([]byte("data"))
updated := crc32.Update(0, table, []byte("data"))
__p(fmt.Sprint(direct == updated)) 
__check("true")
}
