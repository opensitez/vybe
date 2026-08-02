// vybe-test: go/hash_crc32_adler32/adler32_differs_from_fnv32_same_input
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs

package main
import "fmt"
import "hash/adler32"
import "hash/fnv"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { data := []byte("test")
a := adler32.Checksum(data)
h := fnv.New32a()
h.Write(data)
__check(fmt.Sprint(uint32(a) != h.Sum32()), "true") }
