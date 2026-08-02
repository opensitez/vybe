// vybe-test: go/hash_crc32_adler32/fnv_new32_write_sum
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
h.Write([]byte("go"))
__check(fmt.Sprint(h.Sum32()), "1786192775") }
