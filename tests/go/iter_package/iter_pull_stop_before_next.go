// vybe-test: go/iter_package/iter_pull_stop_before_next
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

func main() { ran := 0
seq := func(yield func(int) bool) { ran++
yield(1)
yield(2) }
next, stop := iter.Pull(seq)
stop()
__check(fmt.Sprint(ran), "0") }
