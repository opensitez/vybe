// vybe-test: go/json_marshal/marshal_nested_object
// origin: languages/go/tests/go/test_json_marshal.rs

package main
import "fmt"
import "encoding/json"
type Child struct { N int }
type Parent struct { Child Child
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

func main() { b, _ := json.Marshal(Parent{Child: Child{N: 1}, Tag: "x"})
__p(fmt.Sprint(string(b))) 
__check("{\"Child\":{\"N\":1},\"Tag\":\"x\"}")
}
