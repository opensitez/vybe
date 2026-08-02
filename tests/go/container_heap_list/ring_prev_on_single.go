// vybe-test: go/container_heap_list/ring_prev_on_single
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

func main() { r := ring.New(1)
r.Value = 9
__check(fmt.Sprint(r.Prev().Value), "9") }
