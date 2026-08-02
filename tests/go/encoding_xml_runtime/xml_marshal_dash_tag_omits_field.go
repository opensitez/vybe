// vybe-test: go/encoding_xml_runtime/xml_marshal_dash_tag_omits_field
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { Hidden string `xml:"-"`
Pub int `xml:"pub"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := xml.Marshal(T{Hidden: "secret", Pub: 2})
__check(fmt.Sprint(string(b)), "<T><pub>2</pub></T>") }
