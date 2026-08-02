// vybe-test: go/container_list_ring/list_push_front_list
// origin: languages/go/tests/go/test_container_list_ring.rs
// vybe-test-mode: compile

package main
import "container/list"
func main() { a := list.New()
b := list.New()
a.PushFrontList(b) }
