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

func main() { value := &buf{}
var rw readWriter = value
rw.write(9)
__p(fmt.Sprint(rw.read())) 
__check("9")
}
