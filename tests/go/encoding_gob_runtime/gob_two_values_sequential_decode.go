// vybe-test: go/encoding_gob_runtime/gob_two_values_sequential_decode
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
enc.Encode(1)
enc.Encode(2)
dec := gob.NewDecoder(&buf)
var a, b int
dec.Decode(&a)
dec.Decode(&b)
__check(fmt.Sprint(a), "1")
__check(fmt.Sprint(b), "2") }
