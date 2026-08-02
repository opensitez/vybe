// vybe-test: go/encoding_xml_runtime/xml_marshal_string_attribute
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { S string `xml:"label,attr"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := xml.Marshal(T{S: "vybe"})
__check(fmt.Sprint(string(b)), "<T label=\"vybe\"></T>") }
