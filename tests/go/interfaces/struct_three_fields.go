// vybe-test: go/interfaces/struct_three_fields
// origin: languages/go/tests/go/test_interfaces.rs

package main
import "fmt"
type Employee struct { Name string
Dept string
Salary int } var __buf string

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

func main() { e := Employee{Name: "Alice", Dept: "Eng", Salary: 90000}
__p(fmt.Sprint(e.Name))
__p(fmt.Sprint(e.Dept))
__p(fmt.Sprint(e.Salary))
__check("Alice\nEng\n90000")
}
