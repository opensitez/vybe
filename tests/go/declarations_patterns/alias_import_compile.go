// vybe-test: go/declarations_patterns/alias_import_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
import f "fmt"
func main() { _ = f.Sprintf("%s", "alias") }
