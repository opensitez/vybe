// vybe-test: go/encoding_gob_runtime/gob_map_string_to_struct
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type Rec struct { Val int }
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

func main() { orig := map[string]Rec{"k": {Val: 9}}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back map[string]Rec
gob.NewDecoder(&buf).Decode(&back)
__p(fmt.Sprint(back["k"].Val)) 
__check("9")
}
