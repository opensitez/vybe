// vybe-test: go/encoding_gob_runtime/gob_encode_decode_int_slice
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

func main() { orig := []int{1, 2, 3}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back []int
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(len(back)), "3")
__check(fmt.Sprint(back[2]), "3") }
