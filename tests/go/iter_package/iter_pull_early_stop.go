// vybe-test: go/iter_package/iter_pull_early_stop
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

func main() { seq := func(yield func(int) bool) { if !yield(10) { return }
yield(20) }
next, stop := iter.Pull(seq)
defer stop()
v, ok := next()
stop()
__check(fmt.Sprint(v), "10")
__check(fmt.Sprint(ok), "true") }
