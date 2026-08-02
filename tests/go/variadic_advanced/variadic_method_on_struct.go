// vybe-test: go/variadic_advanced/variadic_method_on_struct
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
type Tally struct{}
func (t Tally) Add(nums ...int) int { s := 0
for _, n := range nums { s += n }
return s }
func main() { fmt.Println(Tally{}.Add(1, 2, 3)) }
