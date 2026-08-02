// vybe-test: go/hash_crc32_adler32/crc32_different_inputs_different_sums
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs

package main
import "fmt"
import "hash/crc32"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := crc32.ChecksumIEEE([]byte("a"))
b := crc32.ChecksumIEEE([]byte("b"))
__check(fmt.Sprint(a != b), "true") }
