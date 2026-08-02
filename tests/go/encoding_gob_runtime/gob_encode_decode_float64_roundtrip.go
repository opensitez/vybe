// vybe-test: go/encoding_gob_runtime/gob_encode_decode_float64_roundtrip
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
gob.NewEncoder(&buf).Encode(3.5)
var f float64
gob.NewDecoder(&buf).Decode(&f)
__check(fmt.Sprint(f), "3.5") }
