// vybe-test: go/strings_ops_extended/split_after_n_limits_segments
// origin: languages/go/tests/go/test_strings_ops_extended.rs

package main
import "fmt"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { parts := strings.SplitAfterN("x-y-z", "-", 2)
__check(fmt.Sprint(len(parts)), "2")
__check(fmt.Sprint(parts[0]), "x-")
__check(fmt.Sprint(parts[1]), "y-z") }
