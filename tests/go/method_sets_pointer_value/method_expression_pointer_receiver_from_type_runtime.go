// vybe-test: go/method_sets_pointer_value/method_expression_pointer_receiver_from_type_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type pair struct { a int }
func (p *pair) set(v int) { p.a = v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { target := &pair{}
fn := (*pair).set
fn(target, 6)
__check(fmt.Sprint(target.a), "6") }
