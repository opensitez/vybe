// vybe-test: go/encoding_gob_runtime/gob_struct_exported_fields_roundtrip
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type Pair struct { A int
B string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { orig := Pair{A: 10, B: "go"}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back Pair
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back.A), "10")
__check(fmt.Sprint(back.B), "go") }
