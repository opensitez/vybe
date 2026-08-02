// vybe-test: go/encoding_gob_runtime/gob_struct_zero_values
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type Empty struct { N int
S string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(Empty{})
var back Empty
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back.N), "0")
__check(fmt.Sprint(back.S), "") }
