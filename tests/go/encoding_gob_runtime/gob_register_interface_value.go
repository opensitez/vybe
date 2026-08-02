// vybe-test: go/encoding_gob_runtime/gob_register_interface_value
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type Counter struct { N int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { gob.Register(&Counter{})
orig := &Counter{N: 12}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back *Counter
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back.N), "12") }
