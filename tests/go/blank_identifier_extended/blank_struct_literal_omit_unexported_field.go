// vybe-test: go/blank_identifier_extended/blank_struct_literal_omit_unexported_field
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
type point struct { x int
y int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { p := point{x: 3}
__check(fmt.Sprint(p.x), "3")
__check(fmt.Sprint(p.y), "0") }
