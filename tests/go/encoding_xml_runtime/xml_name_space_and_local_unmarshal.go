// vybe-test: go/encoding_xml_runtime/xml_name_space_and_local_unmarshal
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { N xml.Name `xml:"tag"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var t T
xml.Unmarshal([]byte(`<T><tag xmlns="urn:ex">leaf</tag></T>`), &t)
__check(fmt.Sprint(t.N.Local), "leaf")
__check(fmt.Sprint(t.N.Space), "urn:ex") }
