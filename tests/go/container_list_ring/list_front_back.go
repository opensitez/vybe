// vybe-test: go/container_list_ring/list_front_back
// origin: languages/go/tests/go/test_container_list_ring.rs

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
l.PushFront("a")
l.PushBack("b")
__check(fmt.Sprint(l.Front().Value) + " " + fmt.Sprint(l.Back().Value), "a b") }
