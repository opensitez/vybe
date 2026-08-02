// vybe-test: go/cover_container_heap/heap_push_pointer_receiver
// origin: languages/go/tests/go/test_cover_container_heap.rs
// vybe-test-mode: compile

package main
import "container/heap"
type evt struct { id int }
type q []*evt
func (h *q) Len() int { return len(*h) }
func (h *q) Less(i, j int) bool { return (*h)[i].id < (*h)[j].id }
func (h *q) Swap(i, j int) { (*h)[i], (*h)[j] = (*h)[j], (*h)[i] }
func main() { var h q
heap.Push(&h, &evt{id: 1}) }
