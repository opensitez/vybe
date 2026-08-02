// vybe-test: go/container_heap_list/list_remove_single
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
e := l.PushBack(99)
l.Remove(e)
__check(fmt.Sprint(l.Len()), "0")
__check(fmt.Sprint(l.Front() == nil), "true") }
