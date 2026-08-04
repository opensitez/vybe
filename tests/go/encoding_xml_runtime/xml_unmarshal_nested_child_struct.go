// vybe-test: go/encoding_xml_runtime/xml_unmarshal_nested_child_struct
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type Inner struct { V int `xml:"v"` }
type Outer struct { Inner Inner `xml:"inner"` }
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

func main() { var o Outer
xml.Unmarshal([]byte(`<Outer><inner><v>8</inner></Outer>`), &o)
__p(fmt.Sprint(o.Inner.V)) 
__check("8")
}
