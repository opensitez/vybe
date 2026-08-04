// vybe-test: go/json_unmarshal_advanced/json_unmarshal_string_tag_bool
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type B struct { Ok bool `json:",string"` }
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

func main() { var b B
json.Unmarshal([]byte(`{"Ok":"true"}`), &b)
__p(fmt.Sprint(b.Ok)) 
__check("true")
}
