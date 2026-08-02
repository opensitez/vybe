// vybe-test: go/hash_crc32_adler32/adler32_size_via_hash_interface
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs

package main
import "fmt"
import "hash/adler32"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { h := adler32.New()
__check(fmt.Sprint(h.Size()), "4") }
