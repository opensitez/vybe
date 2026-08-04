// vybe-test: go/encoding_gob_runtime/gob_map_bool_keys_not_supported_use_int
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

func main() { orig := map[int]bool{0: false, 1: true}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back map[int]bool
gob.NewDecoder(&buf).Decode(&back)
__p(fmt.Sprint(back[1])) 
__check("true")
}
