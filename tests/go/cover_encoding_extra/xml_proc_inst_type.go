// vybe-test: go/cover_encoding_extra/xml_proc_inst_type
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/xml"
func main() { p := xml.ProcInst{}
_ = p.Target }
