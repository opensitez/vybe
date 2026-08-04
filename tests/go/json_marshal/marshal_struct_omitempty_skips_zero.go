// vybe-test: go/json_marshal/marshal_struct_omitempty_skips_zero
// origin: languages/go/tests/go/test_json_marshal.rs

package main
import "fmt"
import "encoding/json"
type Data struct { Count int `json:",omitempty"`
Label string `json:",omitempty"` }
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

func main() { b, _ := json.Marshal(Data{})
__p(fmt.Sprint(string(b))) 
__check("{}")
}
