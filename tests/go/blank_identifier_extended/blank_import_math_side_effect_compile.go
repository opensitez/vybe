// vybe-test: go/blank_identifier_extended/blank_import_math_side_effect_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
import _ "math"
func main() {}
