// vybe-test: go/container_heap_list/ring_do_mutate
// origin: languages/go/tests/go/test_container_heap_list.rs
// vybe-test-mode: compile

package main
import "container/ring"
func main() { r := ring.New(2)
r.Do(func(v interface{}) {}) }
