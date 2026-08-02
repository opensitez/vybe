// vybe-test: go/container_heap_list/list_push_back_nil_list
// origin: languages/go/tests/go/test_container_heap_list.rs
// vybe-test-mode: compile

package main
import "container/list"
func main() { a := list.New()
a.PushBackList(list.New()) }
