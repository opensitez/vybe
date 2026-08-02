// vybe-test: go/function_literals_closures/closure_with_blank_import_capture_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { fn := func() { fmt.Println(1) }
fn() }
