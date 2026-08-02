// vybe-test: go/variadic_advanced/spread_string_runes_as_string_variadic
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func concat(parts ...string) string { out := ""
for _, p := range parts { out += p }
return out }
func main() { letters := []string{"a", "b"}
fmt.Println(concat(letters...)) }
