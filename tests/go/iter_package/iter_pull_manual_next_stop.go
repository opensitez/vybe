// vybe-test: go/iter_package/iter_pull_manual_next_stop
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

func main() { seq := func(yield func(int) bool) { yield(1)
yield(2)
yield(3) }
next, stop := iter.Pull(seq)
defer stop()
v1, ok1 := next()
v2, ok2 := next()
_, ok3 := next()
_, ok4 := next()
__check(fmt.Sprint(v1), "1")
__check(fmt.Sprint(v2), "2")
__check(fmt.Sprint(ok1 && ok2 && ok3 && !ok4), "true") }
