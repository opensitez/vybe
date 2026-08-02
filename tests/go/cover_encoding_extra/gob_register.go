// vybe-test: go/cover_encoding_extra/gob_register
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
func main() { gob.Register(int(0)) }
