// vybe-test: go/container_list_ring/list_move_before
// origin: languages/go/tests/go/test_container_list_ring.rs
// vybe-test-mode: compile

package main
import "container/list"
func main() { l := list.New()
a := l.PushBack(1)
b := l.PushBack(2)
l.MoveBefore(b, a) }
