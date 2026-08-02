// vybe-test: go/container_heap_list/ring_new_zero
// origin: languages/go/tests/go/test_container_heap_list.rs
// vybe-test-mode: compile

package main
import "container/ring"
func main() { r := ring.New(0)
_ = r.Next() }
