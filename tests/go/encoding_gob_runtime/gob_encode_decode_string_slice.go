// vybe-test: go/encoding_gob_runtime/gob_encode_decode_string_slice
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

func main() { orig := []string{"a", "b"}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back []string
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back[0]), "a")
__check(fmt.Sprint(back[1]), "b") }
