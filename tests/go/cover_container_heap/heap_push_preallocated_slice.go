// vybe-test: go/cover_container_heap/heap_push_preallocated_slice
// origin: languages/go/tests/go/test_cover_container_heap.rs
// vybe-test-mode: compile

package main
import "container/heap"
type item struct { pri int
label string }
type pq struct { items []item }
func (h pq) Len() int { return len(h.items) }
func (h pq) Less(i, j int) bool { return h.items[i].pri < h.items[j].pri }
func (h pq) Swap(i, j int) { h.items[i], h.items[j] = h.items[j], h.items[i] }
func main() { h := &pq{items: make([]item, 0, 4)}
heap.Push(h, item{pri: 7}) }
