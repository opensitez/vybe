// vybe-test: go/encoding_xml_runtime/xml_marshal_nested_child_struct
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type Inner struct { V int `xml:"v"` }
type Outer struct { Inner Inner `xml:"inner"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := xml.Marshal(Outer{Inner: Inner{V: 3}})
__check(fmt.Sprint(string(b)), "<Outer><inner><v>3</v></inner></Outer>") }
