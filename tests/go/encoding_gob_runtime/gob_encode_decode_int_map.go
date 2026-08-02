// vybe-test: go/encoding_gob_runtime/gob_encode_decode_int_map
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

func main() { orig := map[string]int{"x": 7, "y": 8}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back map[string]int
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back["x"]), "7")
__check(fmt.Sprint(back["y"]), "8") }
