// vybe-test: go/cover_container_heap/heap_fix_equal_priority
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
func main() { h := &pq{items: []item{{pri: 1}, {pri: 1}}}
heap.Init(h)
heap.Fix(h, 1) }
