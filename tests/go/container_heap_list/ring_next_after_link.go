// vybe-test: go/container_heap_list/ring_next_after_link
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

func main() { a := ring.New(1)
b := ring.New(1)
a.Value = 10
b.Value = 20
a.Link(b)
__check(fmt.Sprint(a.Next().Value), "20") }
