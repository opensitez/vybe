// vybe-test: go/variadic_advanced/variadic_string_max_length
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func longest(words ...string) int { m := 0
for _, w := range words { if len(w) > m { m = len(w) } }
return m }
func main() { fmt.Println(longest("go", "vybe", "a")) }
