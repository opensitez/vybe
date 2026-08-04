// vybe-test: go/encoding_xml_runtime/xml_unmarshal_float_field
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { F float64 `xml:"f"` }
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

func main() { var t T
xml.Unmarshal([]byte(`<T><f>3.14</f></T>`), &t)
__p(fmt.Sprint(t.F)) 
__check("3.14")
}
