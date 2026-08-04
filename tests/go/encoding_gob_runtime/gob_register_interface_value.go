// vybe-test: go/encoding_gob_runtime/gob_register_interface_value
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type Counter struct { N int }
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

func main() { gob.Register(&Counter{})
orig := &Counter{N: 12}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back *Counter
gob.NewDecoder(&buf).Decode(&back)
__p(fmt.Sprint(back.N)) 
__check("12")
}
