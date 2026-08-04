// vybe-test: go/encoding_xml_runtime/xml_marshal_indent_adds_newlines
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { N int `xml:"n"` }
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

func main() { b, _ := xml.MarshalIndent(T{N: 1}, "", "  ")
__p(fmt.Sprint(len(b) > len([]byte("<T><n>1</n></T>")))) 
__check("true")
}
