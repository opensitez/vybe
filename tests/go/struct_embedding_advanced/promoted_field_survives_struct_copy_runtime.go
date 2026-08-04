// vybe-test: go/struct_embedding_advanced/promoted_field_survives_struct_copy_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

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

func main() { original := outer{inner: inner{count: 2}}
copy := original
copy.count = 5
__p(fmt.Sprint(original.count))
__p(fmt.Sprint(copy.count))
__check("2\n5")
}
