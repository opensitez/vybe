// vybe-test: go/pointers/pointer_array_modify
// origin: languages/go/tests/go/test_pointers.rs

package main
import "fmt"
func modify(arr *[3]int) { arr[0] = 99 }
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

func main() { a := [3]int{1, 2, 3}
modify(&a)
__p(fmt.Sprint(a[0]))
__check("99")
}
