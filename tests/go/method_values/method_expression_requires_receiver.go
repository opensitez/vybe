// vybe-test: go/method_values/method_expression_requires_receiver
// origin: languages/go/tests/go/test_method_values.rs

package main
import "fmt"
type box struct { v int }
func (b box) get() int { return b.v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { f := box.get
__check(fmt.Sprint(f(box{v:9})), "9") }
