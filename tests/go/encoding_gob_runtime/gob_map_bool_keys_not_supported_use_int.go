// vybe-test: go/encoding_gob_runtime/gob_map_bool_keys_not_supported_use_int
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

func main() { orig := map[int]bool{0: false, 1: true}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back map[int]bool
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back[1]), "true") }
