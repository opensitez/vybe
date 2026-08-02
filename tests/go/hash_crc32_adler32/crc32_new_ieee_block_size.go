// vybe-test: go/hash_crc32_adler32/crc32_new_ieee_block_size
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

func main() { h := crc32.NewIEEE()
__check(fmt.Sprint(h.BlockSize()), "1") }
