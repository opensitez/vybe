// vybe-test: go/json_unmarshal_advanced/json_unmarshal_embedded_anonymous_tag
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type Meta struct { Ver int `json:"ver"` }
type Doc struct { Meta
Title string }
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

func main() { var d Doc
json.Unmarshal([]byte(`{"ver":2,"Title":"x"}`), &d)
__p(fmt.Sprint(d.Ver))
__p(fmt.Sprint(d.Title)) 
__check("2\nx")
}
