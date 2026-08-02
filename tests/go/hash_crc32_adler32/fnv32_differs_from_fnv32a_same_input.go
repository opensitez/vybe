// vybe-test: go/hash_crc32_adler32/fnv32_differs_from_fnv32a_same_input
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs

package main
import "fmt"
import "hash/fnv"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { data := []byte("abc")
h1 := fnv.New32()
h1.Write(data)
h2 := fnv.New32a()
h2.Write(data)
__check(fmt.Sprint(h1.Sum32() != h2.Sum32()), "true") }
