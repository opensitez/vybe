// vybe-test: go/hash_crc32_adler32/fnv_sum_append_nil
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs

package main
import "fmt"
import "hash/fnv"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { h := fnv.New32()
h.Write([]byte("x"))
sum := h.Sum(nil)
__check(fmt.Sprint(len(sum)), "4") }
