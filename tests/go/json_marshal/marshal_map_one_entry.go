// vybe-test: go/json_marshal/marshal_map_one_entry
// origin: languages/go/tests/go/test_json_marshal.rs

package main
import "fmt"
import "encoding/json"
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

func main() { b, _ := json.Marshal(map[string]int{"a": 1})
__p(fmt.Sprint(string(b))) 
__check("{\"a\":1}")
}
