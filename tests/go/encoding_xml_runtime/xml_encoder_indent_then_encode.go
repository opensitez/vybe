// vybe-test: go/encoding_xml_runtime/xml_encoder_indent_then_encode
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
import "bytes"
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

func main() { buf := bytes.NewBuffer(nil)
e := xml.NewEncoder(buf)
e.Indent("", "  ")
e.Encode(T{N: 2})
__p(fmt.Sprint(buf.Len() > 0)) 
__check("true")
}
