// vybe-test: go/variadic_advanced/forward_variadic_triple_hop
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func end(words ...string) int { return len(words) }
func mid(words ...string) int { return end(words...) }
func start(words ...string) int { return mid(words...) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(start("a", "b")), "2") }
