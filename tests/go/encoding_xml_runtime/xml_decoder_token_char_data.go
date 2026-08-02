// vybe-test: go/encoding_xml_runtime/xml_decoder_token_char_data
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

func main() { d := xml.NewDecoder(strings.NewReader(`<a>xy</a>`))
d.Token()
tok, _ := d.Token()
cd := tok.(xml.CharData)
__check(fmt.Sprint(string(cd)), "xy") }
