// vybe-test: go/encoding_xml_runtime/xml_unmarshal_nested_child_struct
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

func main() { var o Outer
xml.Unmarshal([]byte(`<Outer><inner><v>8</inner></Outer>`), &o)
__check(fmt.Sprint(o.Inner.V), "8") }
