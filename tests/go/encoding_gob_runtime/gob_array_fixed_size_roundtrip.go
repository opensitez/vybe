// vybe-test: go/encoding_gob_runtime/gob_array_fixed_size_roundtrip
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { orig := [3]int{4, 5, 6}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back [3]int
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back[0]), "4")
__check(fmt.Sprint(back[2]), "6") }
