// vybe-test: go/cover_container_heap/heap_init_custom_order
// origin: languages/go/tests/go/test_cover_container_heap.rs
// vybe-test-mode: compile

package main
import "container/heap"
type task struct { deadline int }
type tq struct { list []task }
func (t tq) Len() int { return len(t.list) }
func (t tq) Less(i, j int) bool { return t.list[i].deadline < t.list[j].deadline }
func (t tq) Swap(i, j int) { t.list[i], t.list[j] = t.list[j], t.list[i] }
func main() { q := &tq{list: []task{{deadline: 10}, {deadline: 5}}}
heap.Init(q) }
