// vybe-test: go/lang_declarations_types/empty_struct_size_compile
// origin: languages/go/tests/go/test_lang_declarations_types.rs
// vybe-test-mode: compile

package main
import "unsafe"
type E struct{}
func main() { _ = unsafe.Sizeof(E{}) }
