// vybe-test: go/hash_crc32_adler32/crc32_differs_from_fnv32_same_input
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs

package main
import "fmt"
import "hash/crc32"
import "hash/fnv"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { data := []byte("test")
c := crc32.ChecksumIEEE(data)
h := fnv.New32()
h.Write(data)
__check(fmt.Sprint(c != h.Sum32()), "true") }
