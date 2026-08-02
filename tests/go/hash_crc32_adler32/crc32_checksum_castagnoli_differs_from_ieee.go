// vybe-test: go/hash_crc32_adler32/crc32_checksum_castagnoli_differs_from_ieee
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

func main() { data := []byte("go")
ieee := crc32.ChecksumIEEE(data)
c := crc32.Checksum(data, crc32.MakeTable(crc32.Castagnoli))
__check(fmt.Sprint(ieee != c), "true") }
