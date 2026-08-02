// vybe-test: go/hash_crc32_adler32/fnv32_differs_from_fnv64_same_input
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

func main() { data := []byte("go")
h32 := fnv.New32()
h32.Write(data)
h64 := fnv.New64()
h64.Write(data)
__check(fmt.Sprint(h32.Sum32() != uint32(h64.Sum64())), "true") }
