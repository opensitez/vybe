// vybe-test: go/pointer_receivers_advanced/new_vs_literal_same_method_result_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type score struct { points int }
func (s *score) double() { s.points = s.points * 2 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := new(score)
a.points = 5
a.double()
b := &score{points: 5}
b.double()
__check(fmt.Sprint(a.points), "10")
__check(fmt.Sprint(b.points), "10")
}
