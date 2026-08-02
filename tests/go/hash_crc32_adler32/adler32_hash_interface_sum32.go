// vybe-test: go/hash_crc32_adler32/adler32_hash_interface_sum32
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs

package main
import "fmt"
import "hash/adler32"
import "hash"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var h hash.Hash32 = adler32.New()
h.Write([]byte("go"))
__check(fmt.Sprint(h.Sum32()), "20906199") }
