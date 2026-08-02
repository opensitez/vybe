// vybe-test: go/stdlib_encoding_misc/xml_unmarshal
// origin: languages/go/tests/go/test_stdlib_encoding_misc.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
type T struct { X int `xml:"x"` }
func main() { var t T
_ = xml.Unmarshal([]byte(`<T x="1"/>`), &t) }
