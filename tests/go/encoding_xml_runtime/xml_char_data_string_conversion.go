// vybe-test: go/encoding_xml_runtime/xml_char_data_string_conversion
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

func main() { cd := xml.CharData([]byte("payload"))
__check(fmt.Sprint(string(cd)), "payload") }
