// vybe-test: go/container_heap_list/ring_do_empty_len_zero
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

func main() { r := ring.New(0)
count := 0
r.Do(func(v interface{}) { count++ })
__check(fmt.Sprint(count), "0") }
