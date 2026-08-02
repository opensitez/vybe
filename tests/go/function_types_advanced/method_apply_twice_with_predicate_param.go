// vybe-test: go/function_types_advanced/method_apply_twice_with_predicate_param
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type tally struct { count int }
func (t *tally) whilePositive(ok func(int) bool) { for ok(t.count) { t.count-- } }
func main() { value := tally{count: 3}
value.whilePositive(func(v int) bool { return v > 0 })
fmt.Println(value.count) }
