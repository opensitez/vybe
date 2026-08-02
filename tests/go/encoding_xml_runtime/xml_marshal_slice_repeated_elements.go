// vybe-test: go/encoding_xml_runtime/xml_marshal_slice_repeated_elements
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

func main() { b, _ := xml.Marshal(T{Items: []int{1, 2}})
__check(fmt.Sprint(string(b)), "<T><item>1</item><item>2</item></T>") }
