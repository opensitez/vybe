// vybe-test: go/container_heap_list/ring_link_combines
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

func main() { a := ring.New(2)
b := ring.New(2)
a.Value = 1
a.Next().Value = 2
b.Value = 3
b.Next().Value = 4
a.Link(b)
__check(fmt.Sprint(a.Len()), "4") }
