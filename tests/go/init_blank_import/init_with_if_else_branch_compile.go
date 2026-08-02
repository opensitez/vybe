// vybe-test: go/init_blank_import/init_with_if_else_branch_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
var flag bool
func init() { if 2 > 1 { flag = true } else { flag = false } }
func main() { _ = flag }
