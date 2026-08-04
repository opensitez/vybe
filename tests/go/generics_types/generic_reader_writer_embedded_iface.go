// vybe-test: go/generics_types/generic_reader_writer_embedded_iface
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Reader[T any] interface { Read() T }
type Writer[T any] interface { Write(T) }
type ReadWriter[T any] interface { Reader[T]
Writer[T] }
type Buffer[T any] struct { data []T }
func (b *Buffer[T]) Read() T { v := b.data[0]
b.data = b.data[1:]
return v }
func (b *Buffer[T]) Write(v T) { b.data = append(b.data, v) }
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

func main() { var rw ReadWriter[int] = &Buffer[int]{}
rw.Write(3)
__p(fmt.Sprint(rw.Read())) 
__check("3")
}
