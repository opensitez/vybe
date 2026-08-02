// vybe-test: go/hash_crc32_adler32/crc32_checksum_ieee_single_byte
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

func main() { __check(fmt.Sprint(crc32.ChecksumIEEE([]byte("a"))), "3904355907") }
