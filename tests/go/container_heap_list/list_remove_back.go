// vybe-test: go/container_heap_list/list_remove_back
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
l.PushFront(1)
e := l.PushBack(2)
l.Remove(e)
__check(fmt.Sprint(l.Back().Value), "1") }
