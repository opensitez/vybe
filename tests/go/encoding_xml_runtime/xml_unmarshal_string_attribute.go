// vybe-test: go/encoding_xml_runtime/xml_unmarshal_string_attribute
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { S string `xml:"name,attr"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var t T
xml.Unmarshal([]byte(`<T name="alice"/>`), &t)
__check(fmt.Sprint(t.S), "alice") }
