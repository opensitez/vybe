// vybe-test: go/variadic_advanced/variadic_with_blank_import_call_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
import "fmt"
func show(parts ...interface{}) { _ = fmt.Sprint(parts...) }
func main() { show(1, 2) }
