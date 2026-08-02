// vybe-test: go/function_literals_closures/closure_capture_loop_var_with_param
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func main() { sum := 0
for i := 1; i <= 3; i++ { func(n int) { sum += n }(i) }
fmt.Println(sum) }
