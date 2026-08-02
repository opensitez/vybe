// vybe-test: go/iter_package/iter_pull2_second_pair
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

func main() { seq := func(yield func(string, int) bool) { yield("a", 1)
yield("b", 2) }
next, stop := iter.Pull2(seq)
defer stop()
next()
k, v, ok := next()
__check(fmt.Sprint(k), "b")
__check(fmt.Sprint(v), "2")
__check(fmt.Sprint(ok), "true") }
