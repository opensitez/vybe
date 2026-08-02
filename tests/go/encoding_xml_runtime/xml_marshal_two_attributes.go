// vybe-test: go/encoding_xml_runtime/xml_marshal_two_attributes
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { A int `xml:"a,attr"`
B string `xml:"b,attr"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := xml.Marshal(T{A: 1, B: "z"})
s := string(b)
__check(fmt.Sprint(len(s) > 10), "true") }
