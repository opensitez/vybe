// vybe-test: go/strings_ops_extended/split_n_unlimited_negative_one
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

func main() { parts := strings.SplitN("one:two:three", ":", -1)
__check(fmt.Sprint(len(parts)), "3") }
