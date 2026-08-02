// vybe-test: go/hash_crc32_adler32/crc32_long_input_nonzero
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs

package main
import "fmt"
import "hash/crc32"
func main() { data := make([]byte, 256)
for i := range data { data[i] = byte(i) }
fmt.Println(crc32.ChecksumIEEE(data) != 0) }
