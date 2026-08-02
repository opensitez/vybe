// vybe-test: go/cover_encoding_extra/xml_unescape
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
func main() { b, _ := xml.Unescape([]byte("&lt;a&gt;"))
_ = b }
