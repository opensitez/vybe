// vybe-test: go/encoding_gob_runtime/gob_encode_decode_int_roundtrip
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

func main() { var buf bytes.Buffer
enc := gob.NewEncoder(&buf)
enc.Encode(42)
var v int
gob.NewDecoder(&buf).Decode(&v)
__check(fmt.Sprint(v), "42") }
