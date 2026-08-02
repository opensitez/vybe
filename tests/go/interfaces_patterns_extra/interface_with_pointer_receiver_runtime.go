// vybe-test: go/interfaces_patterns_extra/interface_with_pointer_receiver_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
type sized interface { size() int }
type box struct { n int }
func (b *box) size() int { return b.n }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value sized = &box{n: 6}
__check(fmt.Sprint(value.size()), "6")
}
