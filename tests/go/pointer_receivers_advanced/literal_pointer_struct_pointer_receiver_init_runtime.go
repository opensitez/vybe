// vybe-test: go/pointer_receivers_advanced/literal_pointer_struct_pointer_receiver_init_runtime
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

func main() { value := &tally{sum: 1}
value.add(4)
__check(fmt.Sprint(value.sum), "5")
}
