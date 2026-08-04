// vybe-test: go/structs/struct_method
// origin: languages/go/tests/go/test_structs.rs

package main
import "fmt"
type Person struct { Name string
Age int } func (p Person) Greet() { __p(fmt.Sprint(p.Name))
} var __buf string

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

func main() { p := Person{Name: "Bob", Age: 25}
p.Greet()
__check("Bob")
}
