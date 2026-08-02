// vybe-test: go/encoding_xml_runtime/xml_encoder_indent_then_encode
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
import "bytes"
type T struct { N int `xml:"n"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { buf := bytes.NewBuffer(nil)
e := xml.NewEncoder(buf)
e.Indent("", "  ")
e.Encode(T{N: 2})
__check(fmt.Sprint(buf.Len() > 0), "true") }
