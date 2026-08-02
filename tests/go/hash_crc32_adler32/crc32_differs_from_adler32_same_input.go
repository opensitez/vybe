// vybe-test: go/hash_crc32_adler32/crc32_differs_from_adler32_same_input
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs

package main
import "fmt"
import "hash/crc32"
import "hash/adler32"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { data := []byte("test")
c := crc32.ChecksumIEEE(data)
a := adler32.Checksum(data)
__check(fmt.Sprint(c != uint32(a)), "true") }
