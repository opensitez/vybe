// vybe-test: go/interface_embedding_methods/reader_writer_write_promoted_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type reader interface { read() int }
type writer interface { write(int) }
type readWriter interface { reader
writer }
type buf struct { data int }
func (b *buf) read() int { return b.data }
func (b *buf) write(n int) { b.data = n }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := &buf{}
var rw readWriter = value
rw.write(9)
__check(fmt.Sprint(rw.read()), "9") }
