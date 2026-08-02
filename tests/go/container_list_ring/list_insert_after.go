// vybe-test: go/container_list_ring/list_insert_after
// origin: languages/go/tests/go/test_container_list_ring.rs
// vybe-test-mode: compile

package main
import "container/list"
func main() { l := list.New()
e := l.PushBack(1)
l.InsertAfter(2, e) }
