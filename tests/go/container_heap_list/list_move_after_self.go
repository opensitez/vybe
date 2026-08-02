// vybe-test: go/container_heap_list/list_move_after_self
// origin: languages/go/tests/go/test_container_heap_list.rs
// vybe-test-mode: compile

package main
import "container/list"
func main() { l := list.New()
e := l.PushBack(1)
l.MoveAfter(e, e) }
