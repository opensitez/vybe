// vybe-test: go/container_heap_list/list_insert_before_root
// origin: languages/go/tests/go/test_container_heap_list.rs
// vybe-test-mode: compile

package main
import "container/list"
func main() { l := list.New()
e := l.PushBack(1)
l.InsertBefore(0, e) }
