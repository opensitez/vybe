// vybe-test: go/function_literals_closures/closure_in_for_range
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func main() { sum := 0
for _, v := range []int{1, 2, 3} { func(x int) { sum += x }(v) }
fmt.Println(sum) }
