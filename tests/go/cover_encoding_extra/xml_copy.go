// vybe-test: go/cover_encoding_extra/xml_copy
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
import "bytes"
func main() { dst := bytes.NewBuffer(nil)
_ = xml.Copy(dst, bytes.NewBufferString("<a/>")) }
