// vybe-test: go/container_list_ring/list_push_back_len
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
l.PushBack(1)
l.PushBack(2)
__check(fmt.Sprint(l.Len()), "2") }
