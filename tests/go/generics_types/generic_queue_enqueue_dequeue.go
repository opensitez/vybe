// vybe-test: go/generics_types/generic_queue_enqueue_dequeue
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Queue[T any] struct { items []T }
func (q *Queue[T]) Enqueue(v T) { q.items = append(q.items, v) }
func (q *Queue[T]) Dequeue() T { v := q.items[0]
q.items = q.items[1:]
return v }
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { var q Queue[int]
q.Enqueue(10)
q.Enqueue(20)
__p(fmt.Sprint(q.Dequeue()))
__p(fmt.Sprint(q.Dequeue())) 
__check("10\n20")
}
