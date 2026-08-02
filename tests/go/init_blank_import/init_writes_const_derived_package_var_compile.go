// vybe-test: go/init_blank_import/init_writes_const_derived_package_var_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
const base = 4
var doubled int
func init() { doubled = base * 2 }
func main() { _ = doubled }
