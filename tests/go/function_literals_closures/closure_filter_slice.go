// vybe-test: go/function_literals_closures/closure_filter_slice
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func main() { nums := []int{1, 2, 3, 4}
evens := func() []int { out := []int{}
for _, n := range nums { if n%2 == 0 { out = append(out, n) } }
return out }
r := evens()
fmt.Println(len(r))
fmt.Println(r[0]) }
