// vybe-test: go/container_heap_list/ring_next_cycles
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

func main() { r := ring.New(3)
r.Value = 10
r = r.Next()
r.Value = 20
r = r.Next()
r.Value = 30
r = r.Next()
__check(fmt.Sprint(r.Value), "10") }
