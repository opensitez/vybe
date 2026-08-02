// vybe-test: go/functions_patterns_extra/variadic_forwarding_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func sum(values ...int) int { total := 0
for _, v := range values { total += v }
return total }
func wrap(values ...int) int { return sum(values...) }
func main() { fmt.Println(wrap(1, 2, 3))
}
