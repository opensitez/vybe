// vybe-test: go/container_heap_list/ring_unlink_splits
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

func main() { r := ring.New(4)
split := r.Unlink(2)
__check(fmt.Sprint(r.Len()), "2")
__check(fmt.Sprint(split.Len()), "2") }
