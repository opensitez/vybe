// vybe-test: go/container_heap_list/ring_do_single_element
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

func main() { r := ring.New(1)
r.Value = 42
count := 0
r.Do(func(v interface{}) { count++; __check(fmt.Sprint(v), "42") })
__check(fmt.Sprint(count), "1") }
