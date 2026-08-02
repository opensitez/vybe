// vybe-test: go/encoding_xml_runtime/xml_unmarshal_pointer_allocates
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { P *int `xml:"p"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var t T
xml.Unmarshal([]byte(`<T><p>11</p></T>`), &t)
__check(fmt.Sprint(*t.P), "11") }
