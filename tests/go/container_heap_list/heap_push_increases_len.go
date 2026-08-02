// vybe-test: go/container_heap_list/heap_push_increases_len
// origin: languages/go/tests/go/test_container_heap_list.rs

package main
import "fmt"
import "container/heap"
type IH []int
func (h IH) Len() int { return len(h) }
func (h IH) Less(i, j int) bool { return h[i] < h[j] }
func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }
func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }
func (h *IH) Pop() interface{} { o := *h
n := len(o)
x := o[n-1]
*h = o[:n-1]
return x }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { h := &IH{}
heap.Init(h)
heap.Push(h, 4)
heap.Push(h, 2)
__check(fmt.Sprint(h.Len()), "2") }
