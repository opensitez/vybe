// vybe-test: go/generics_types/generic_queue_enqueue_dequeue
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Queue[T any] struct { items []T }
func (q *Queue[T]) Enqueue(v T) { q.items = append(q.items, v) }
func (q *Queue[T]) Dequeue() T { v := q.items[0]
q.items = q.items[1:]
return v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var q Queue[int]
q.Enqueue(10)
q.Enqueue(20)
__check(fmt.Sprint(q.Dequeue()), "10")
__check(fmt.Sprint(q.Dequeue()), "20") }
