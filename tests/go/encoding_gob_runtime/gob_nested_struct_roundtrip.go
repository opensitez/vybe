// vybe-test: go/encoding_gob_runtime/gob_nested_struct_roundtrip
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type Inner struct { V int }
type Outer struct { Inner Inner
Tag string }
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { orig := Outer{Inner: Inner{V: 4}, Tag: "t"}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back Outer
gob.NewDecoder(&buf).Decode(&back)
__p(fmt.Sprint(back.Inner.V))
__p(fmt.Sprint(back.Tag)) 
__check("4\nt")
}
