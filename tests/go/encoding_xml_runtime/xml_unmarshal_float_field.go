// vybe-test: go/encoding_xml_runtime/xml_unmarshal_float_field
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

func main() { var t T
xml.Unmarshal([]byte(`<T><f>3.14</f></T>`), &t)
__check(fmt.Sprint(t.F), "3.14") }
