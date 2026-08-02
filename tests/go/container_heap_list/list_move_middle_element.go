// vybe-test: go/container_heap_list/list_move_middle_element
// origin: languages/go/tests/go/test_container_heap_list.rs

package main
import "fmt"
import "container/list"
func main() { l := list.New()
a := l.PushBack(1)
b := l.PushBack(2)
c := l.PushBack(3)
l.MoveBefore(b, c)
var vals []int
for e := l.Front(); e != nil; e = e.Next() { vals = append(vals, e.Value.(int)) }
fmt.Println(vals[0])
fmt.Println(vals[2]) }
