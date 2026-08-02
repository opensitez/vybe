// vybe-test: go/function_literals_closures/closure_string_builder_pattern
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func main() { parts := []string{"go", "lang"}
join := func(sep string) string { s := ""
for i, p := range parts { if i > 0 { s += sep }
s += p }
return s }
fmt.Println(join("-")) }
