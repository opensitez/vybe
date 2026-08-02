// vybe-test: go/container_list_ring/list_push_front
// origin: languages/go/tests/go/test_container_list_ring.rs
// vybe-test-mode: compile

package main
import "container/list"
func main() { l := list.New()
l.PushFront(1) }
