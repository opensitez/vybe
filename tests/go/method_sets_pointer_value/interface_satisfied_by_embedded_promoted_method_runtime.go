// vybe-test: go/method_sets_pointer_value/interface_satisfied_by_embedded_promoted_method_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type runner interface { run() int }
type legs struct{}
func (legs) run() int { return 42 }
type athlete struct { legs }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var r runner = athlete{}
__check(fmt.Sprint(r.run()), "42") }
