// vybe-test: go/declarations_patterns/exported_identifier_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
type Person struct { Name string }
func main() { _ = Person{Name: "Ada"} }
