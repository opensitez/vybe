// vybe-test: go/container_heap_list/list_insert_after
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
e := l.PushBack(1)
l.InsertAfter(2, e)
__check(fmt.Sprint(l.Front().Value), "1")
__check(fmt.Sprint(l.Back().Value), "2") }
