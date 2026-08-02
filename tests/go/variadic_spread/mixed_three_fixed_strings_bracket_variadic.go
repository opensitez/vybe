// vybe-test: go/variadic_spread/mixed_three_fixed_strings_bracket_variadic
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func bracket(open string, close string, sep string, parts ...string) string { out := open
for i, p := range parts { if i > 0 { out += sep }
out += p }
return out + close }
func main() { fmt.Println(bracket("[", "]", ",", "a", "b"))
}
