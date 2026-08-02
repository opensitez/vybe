// vybe-test: go/hash_crc32_adler32/crc32_update_matches_checksum
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

func main() { table := crc32.IEEETable
direct := crc32.ChecksumIEEE([]byte("data"))
updated := crc32.Update(0, table, []byte("data"))
__check(fmt.Sprint(direct == updated), "true") }
