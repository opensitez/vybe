// vybe-test: go/encoding_xml_runtime/xml_decoder_token_end_element
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

func main() { d := xml.NewDecoder(strings.NewReader(`<a/>`))
d.Token()
tok, _ := d.Token()
end := tok.(xml.EndElement)
__check(fmt.Sprint(end.Name.Local), "a") }
