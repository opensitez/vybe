// vybe-test: go/encoding_xml_runtime/xml_marshal_nested_child_struct
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

func main() { b, _ := xml.Marshal(Outer{Inner: Inner{V: 3}})
__p(fmt.Sprint(string(b))) 
__check("<Outer><inner><v>3</v></inner></Outer>")
}
