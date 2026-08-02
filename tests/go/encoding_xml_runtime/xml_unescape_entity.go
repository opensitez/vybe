// vybe-test: go/encoding_xml_runtime/xml_unescape_entity
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := xml.Unescape([]byte("&lt;tag&gt;"))
__check(fmt.Sprint(string(b)), "<tag>") }
