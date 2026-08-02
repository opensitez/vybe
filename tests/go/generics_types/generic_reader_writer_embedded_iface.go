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
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var rw ReadWriter[int] = &Buffer[int]{}
rw.Write(3)
__check(fmt.Sprint(rw.Read()), "3") }
