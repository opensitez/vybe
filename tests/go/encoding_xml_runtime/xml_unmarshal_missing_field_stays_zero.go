// vybe-test: go/encoding_xml_runtime/xml_unmarshal_missing_field_stays_zero
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { N int `xml:"n"`
S string `xml:"s"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var t T
xml.Unmarshal([]byte(`<T><n>1</n></T>`), &t)
__check(fmt.Sprint(t.N), "1")
__check(fmt.Sprint(t.S), "") }
