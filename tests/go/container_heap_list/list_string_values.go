// vybe-test: go/container_heap_list/list_string_values
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
l.PushBack("go")
l.PushBack("vybe")
__check(fmt.Sprint(l.Front().Value), "go")
__check(fmt.Sprint(l.Back().Value), "vybe") }
