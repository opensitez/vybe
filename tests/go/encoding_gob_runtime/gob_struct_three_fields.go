// vybe-test: go/encoding_gob_runtime/gob_struct_three_fields
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type Triple struct { A int
B int
C string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { orig := Triple{A: 1, B: 2, C: "c"}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back Triple
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back.B), "2")
__check(fmt.Sprint(back.C), "c") }
