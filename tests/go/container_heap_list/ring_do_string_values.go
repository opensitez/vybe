// vybe-test: go/container_heap_list/ring_do_string_values
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

func main() { r := ring.New(2)
r.Value = "a"
r.Next().Value = "b"
first := ""
r.Do(func(v interface{}) { if first == "" { first = v.(string) } })
__check(fmt.Sprint(first), "a") }
