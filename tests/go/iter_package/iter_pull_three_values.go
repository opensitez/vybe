// vybe-test: go/iter_package/iter_pull_three_values
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "iter"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { seq := func(yield func(int) bool) { yield(5)
yield(6)
yield(7) }
next, stop := iter.Pull(seq)
defer stop()
a, _ := next()
b, _ := next()
c, _ := next()
__check(fmt.Sprint(a + b + c), "18") }
