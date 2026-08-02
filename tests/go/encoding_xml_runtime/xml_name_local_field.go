// vybe-test: go/encoding_xml_runtime/xml_name_local_field
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { N xml.Name `xml:"item"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { t := T{N: xml.Name{Local: "widget"}}
b, _ := xml.Marshal(t)
__check(fmt.Sprint(string(b)), "<T><item>widget</item></T>") }
