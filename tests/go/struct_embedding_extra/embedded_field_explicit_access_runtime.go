// vybe-test: go/struct_embedding_extra/embedded_field_explicit_access_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type inner struct { count int }
type outer struct { inner }
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

func main() { value := outer{inner: inner{count: 7}}
__p(fmt.Sprint(value.inner.count))
__check("7")
}
