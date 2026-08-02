// vybe-test: go/container_heap_list/list_push_front_back_mixed
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
l.PushBack(2)
l.PushFront(1)
l.PushBack(3)
__check(fmt.Sprint(l.Front().Value), "1")
__check(fmt.Sprint(l.Back().Value), "3")
__check(fmt.Sprint(l.Len()), "3") }
