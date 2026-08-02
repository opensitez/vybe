// vybe-test: go/init_blank_import/multiple_blank_imports_different_packages_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
import _ "strings"
import _ "math"
func main() {}
