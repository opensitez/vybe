// vybe-test: go/init_blank_import/four_init_functions_sequential_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
var step int
func init() { step = 1 }
func init() { step++ }
func init() { step++ }
func init() { step++ }
func main() { _ = step }
