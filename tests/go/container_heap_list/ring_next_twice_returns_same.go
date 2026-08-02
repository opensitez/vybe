// vybe-test: go/container_heap_list/ring_next_twice_returns_same
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
r.Value = 7
__check(fmt.Sprint(r.Next().Value), "7") }
