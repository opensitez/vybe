// vybe-test: go/encoding_xml_runtime/xml_marshal_roundtrip_int
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
type T struct { N int `xml:"n,attr"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { orig := T{N: 55}
b, _ := xml.Marshal(orig)
var back T
xml.Unmarshal(b, &back)
__check(fmt.Sprint(back.N), "55") }
