// vybe-test: go/iter_package/iter_pull2_exhausted
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

func main() { seq := func(yield func(int, int) bool) { yield(0, 0) }
next, stop := iter.Pull2(seq)
defer stop()
_, _, ok1 := next()
_, _, ok2 := next()
__check(fmt.Sprint(ok1), "true")
__check(fmt.Sprint(ok2), "false") }
