// vybe-test: go/encoding_gob_runtime/gob_nested_struct_roundtrip
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type Inner struct { V int }
type Outer struct { Inner Inner
Tag string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { orig := Outer{Inner: Inner{V: 4}, Tag: "t"}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back Outer
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back.Inner.V), "4")
__check(fmt.Sprint(back.Tag), "t") }
