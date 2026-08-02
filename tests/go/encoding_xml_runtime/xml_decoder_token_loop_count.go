// vybe-test: go/encoding_xml_runtime/xml_decoder_token_loop_count
// origin: languages/go/tests/go/test_encoding_xml_runtime.rs

package main
import "fmt"
import "encoding/xml"
import "strings"
func main() { d := xml.NewDecoder(strings.NewReader(`<a><b/><c/></a>`))
n := 0
for { tok, err := d.Token()
if err != nil { break }
if _, ok := tok.(xml.StartElement); ok { n++ } }
fmt.Println(n) }
