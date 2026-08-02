// vybe-test: go/encoding_xml_runtime/xml_unmarshal_slice_repeated_elements
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { Items []int `xml:"item"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var t T
xml.Unmarshal([]byte(`<T><item>4</item><item>5</item></T>`), &t)
__check(fmt.Sprint(len(t.Items)), "2")
__check(fmt.Sprint(t.Items[1]), "5") }
