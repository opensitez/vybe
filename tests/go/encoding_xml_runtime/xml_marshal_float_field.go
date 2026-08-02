// vybe-test: go/encoding_xml_runtime/xml_marshal_float_field
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { F float64 `xml:"f"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := xml.Marshal(T{F: 2.5})
__check(fmt.Sprint(string(b)), "<T><f>2.5</f></T>") }
