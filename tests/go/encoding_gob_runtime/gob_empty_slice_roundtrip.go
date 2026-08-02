// vybe-test: go/encoding_gob_runtime/gob_empty_slice_roundtrip
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

func main() { orig := []int{}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back []int
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(len(back)), "0") }
