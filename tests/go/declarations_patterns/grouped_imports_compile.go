// vybe-test: go/declarations_patterns/grouped_imports_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
import ( "fmt"; "strings" )
func main() { _, _ = fmt.Sprintf("%s", "x"), strings.HasPrefix("go", "g") }
