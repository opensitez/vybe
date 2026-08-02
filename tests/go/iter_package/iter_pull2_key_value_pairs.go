// vybe-test: go/iter_package/iter_pull2_key_value_pairs
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

func main() { seq := func(yield func(int, string) bool) { yield(1, "a")
yield(2, "b") }
next, stop := iter.Pull2(seq)
defer stop()
k, v, ok := next()
__check(fmt.Sprint(k), "1")
__check(fmt.Sprint(v), "a")
__check(fmt.Sprint(ok), "true") }
