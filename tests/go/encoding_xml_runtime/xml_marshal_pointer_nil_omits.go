// vybe-test: go/encoding_xml_runtime/xml_marshal_pointer_nil_omits
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { P *int `xml:"p,omitempty"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := xml.Marshal(T{})
__check(fmt.Sprint(string(b)), "<T></T>") }
