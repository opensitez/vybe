// vybe-test: go/container_heap_list/list_move_before_reorders
// origin: languages/go/tests/go/test_container_heap_list.rs

package main
import "fmt"
import "container/list"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { l := list.New()
a := l.PushBack(1)
b := l.PushBack(2)
c := l.PushBack(3)
l.MoveBefore(c, a)
__check(fmt.Sprint(l.Front().Value), "3")
__check(fmt.Sprint(l.Back().Value), "2") }
