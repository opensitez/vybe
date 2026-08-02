// vybe-test: go/encoding_xml_runtime/xml_decoder_token_start_element_local
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { d := xml.NewDecoder(strings.NewReader(`<root><child/></root>`))
tok, _ := d.Token()
start := tok.(xml.StartElement)
__check(fmt.Sprint(start.Name.Local), "root") }
