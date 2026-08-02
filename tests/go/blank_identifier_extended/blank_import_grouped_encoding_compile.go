// vybe-test: go/blank_identifier_extended/blank_import_grouped_encoding_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
import ( _ "encoding/json"; _ "encoding/hex" )
func main() {}
