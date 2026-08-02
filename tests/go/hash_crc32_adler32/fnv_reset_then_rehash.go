// vybe-test: go/hash_crc32_adler32/fnv_reset_then_rehash
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

func main() { h := fnv.New64()
h.Write([]byte("go"))
h.Reset()
h.Write([]byte("go"))
__check(fmt.Sprint(h.Sum64()), "590641186866933191") }
