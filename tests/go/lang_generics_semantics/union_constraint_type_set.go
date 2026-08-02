// vybe-test: go/lang_generics_semantics/union_constraint_type_set
// origin: languages/go/tests/go/test_lang_generics_semantics.rs

package main
import "fmt"
func Describe[T int | string](v T) { fmt.Printf("%v", v) }
func main() { Describe(1) }
