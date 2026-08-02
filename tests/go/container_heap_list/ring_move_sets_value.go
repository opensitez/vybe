// vybe-test: go/container_heap_list/ring_move_sets_value
// origin: languages/go/tests/go/test_container_heap_list.rs

package main
import "fmt"
import "container/ring"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { r := ring.New(3)
r.Value = 0
r.Move(2)
__check(fmt.Sprint(r.Value), "2") }
