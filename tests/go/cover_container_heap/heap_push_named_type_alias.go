// vybe-test: go/cover_container_heap/heap_push_named_type_alias
// origin: languages/go/tests/go/test_cover_container_heap.rs
// vybe-test-mode: compile

package main
import "container/heap"
type rank int
type entry struct { r rank }
type hq struct { data []entry }
func (h hq) Len() int { return len(h.data) }
func (h hq) Less(i, j int) bool { return h.data[i].r < h.data[j].r }
func (h hq) Swap(i, j int) { h.data[i], h.data[j] = h.data[j], h.data[i] }
func main() { h := &hq{}
heap.Push(h, entry{r: 1}) }
