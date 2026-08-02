// vybe-test: go/container_heap_list/list_iterate_forward
// origin: languages/go/tests/go/test_container_heap_list.rs

package main
import "fmt"
import "container/list"
func main() { l := list.New()
l.PushBack(1)
l.PushBack(2)
l.PushBack(3)
sum := 0
for e := l.Front(); e != nil; e = e.Next() { sum += e.Value.(int) }
fmt.Println(sum) }
