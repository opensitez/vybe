// vybe-test: go/encoding_xml_runtime/xml_escape_ampersand_in_text
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
xml.EscapeText(&buf, []byte("a&b"))
__check(fmt.Sprint(buf.String()), "a&amp;b") }
