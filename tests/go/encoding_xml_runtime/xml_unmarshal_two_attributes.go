// vybe-test: go/encoding_xml_runtime/xml_unmarshal_two_attributes
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

func main() { var t T
xml.Unmarshal([]byte(`<T a="2" b="y"/>`), &t)
__check(fmt.Sprint(t.A), "2")
__check(fmt.Sprint(t.B), "y") }
