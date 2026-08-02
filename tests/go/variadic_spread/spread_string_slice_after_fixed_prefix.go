// vybe-test: go/variadic_spread/spread_string_slice_after_fixed_prefix
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func tag(prefix string, words ...string) { for _, w := range words { fmt.Println(prefix + w) } }
func main() { rest := []string{"go", "vybe"}
tag(">", rest...)
}
