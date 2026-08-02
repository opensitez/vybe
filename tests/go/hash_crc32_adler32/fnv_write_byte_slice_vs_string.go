// vybe-test: go/hash_crc32_adler32/fnv_write_byte_slice_vs_string
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

func main() { h1 := fnv.New32a()
h1.Write([]byte("go"))
h2 := fnv.New32a()
h2.Write([]byte("go"))
__check(fmt.Sprint(h1.Sum32() == h2.Sum32()), "true") }
