// vybe-test: go/encoding_xml_runtime/xml_marshal_omitempty_includes_nonzero_int
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { N int `xml:"n,omitempty"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := xml.Marshal(T{N: 1})
__check(fmt.Sprint(string(b)), "<T><n>1</n></T>") }
