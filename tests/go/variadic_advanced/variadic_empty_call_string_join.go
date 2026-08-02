// vybe-test: go/variadic_advanced/variadic_empty_call_string_join
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func join(sep string, parts ...string) string { out := ""
for i, p := range parts { if i > 0 { out += sep }
out += p }
return out }
func main() { fmt.Println(join(",", )) }
