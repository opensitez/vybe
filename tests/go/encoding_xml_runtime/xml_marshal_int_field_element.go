// vybe-test: go/encoding_xml_runtime/xml_marshal_int_field_element
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { N int `xml:"n"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := xml.Marshal(T{N: 7})
__check(fmt.Sprint(string(b)), "<T><n>7</n></T>") }
