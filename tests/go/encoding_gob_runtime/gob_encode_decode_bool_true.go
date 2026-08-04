// vybe-test: go/encoding_gob_runtime/gob_encode_decode_bool_true
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
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

func main() { var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(true)
var v bool
gob.NewDecoder(&buf).Decode(&v)
__p(fmt.Sprint(v)) 
__check("true")
}
