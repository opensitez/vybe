// vybe-test: go/container_heap_list/list_back_nil_when_empty_compile
// origin: languages/go/tests/go/test_container_heap_list.rs
// vybe-test-mode: compile

package main
import "container/list"
func main() { l := list.New()
_ = l.Back() }
