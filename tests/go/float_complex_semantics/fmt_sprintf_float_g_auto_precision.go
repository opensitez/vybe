// vybe-test: go/float_complex_semantics/fmt_sprintf_float_g_auto_precision
// origin: languages/go/tests/go/test_float_complex_semantics.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { _ = fmt.Sprintf("%#g", 3.14) }
