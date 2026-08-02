// vybe-test: go/hash_crc32_adler32/fnv_hash64_interface
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs

package main
import "fmt"
import "hash/fnv"
import "hash"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var h hash.Hash64 = fnv.New64()
h.Write([]byte("go"))
__check(fmt.Sprint(h.Sum64()), "590641186866933191") }
