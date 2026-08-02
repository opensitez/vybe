// vybe-test: go/pointer_receivers_advanced/new_struct_pointer_receiver_zero_init_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type tally struct { sum int }
func (t *tally) add(v int) { t.sum += v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := new(tally)
value.add(3)
value.add(4)
__check(fmt.Sprint(value.sum), "7")
}
