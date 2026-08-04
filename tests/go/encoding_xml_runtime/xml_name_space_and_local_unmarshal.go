// vybe-test: go/encoding_xml_runtime/xml_name_space_and_local_unmarshal
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { N xml.Name `xml:"tag"` }
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
xml.Unmarshal([]byte(`<T><tag xmlns="urn:ex">leaf</tag></T>`), &t)
__p(fmt.Sprint(t.N.Local))
__p(fmt.Sprint(t.N.Space)) 
__check("leaf\nurn:ex")
}
