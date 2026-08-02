// vybe-test: go/encoding_gob_runtime/gob_register_then_encode_custom_type
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type ID struct { N int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { gob.Register(ID{})
orig := ID{N: 3}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back ID
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back.N), "3") }
