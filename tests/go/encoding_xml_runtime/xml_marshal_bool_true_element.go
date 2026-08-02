// vybe-test: go/encoding_xml_runtime/xml_marshal_bool_true_element
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { Ok bool `xml:"ok"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := xml.Marshal(T{Ok: true})
__check(fmt.Sprint(string(b)), "<T><ok>true</ok></T>") }
