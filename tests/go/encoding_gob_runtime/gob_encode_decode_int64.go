// vybe-test: go/encoding_gob_runtime/gob_encode_decode_int64
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
gob.NewEncoder(&buf).Encode(int64(1000000))
var v int64
gob.NewDecoder(&buf).Decode(&v)
__check(fmt.Sprint(v), "1000000") }
