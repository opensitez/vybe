// vybe-test: go/init_blank_import/init_with_for_loop_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
var count int
func init() { for i := 0; i < 3; i++ { count++ } }
func main() { _ = count }
