// vybe-test: go/hash_crc32_adler32/adler32_long_input_nonzero
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs

package main
import "fmt"
import "hash/adler32"
func main() { data := make([]byte, 256)
for i := range data { data[i] = byte(i) }
fmt.Println(adler32.Checksum(data) != 1) }
