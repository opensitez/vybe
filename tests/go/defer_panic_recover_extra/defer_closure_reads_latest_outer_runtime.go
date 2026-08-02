// vybe-test: go/defer_panic_recover_extra/defer_closure_reads_latest_outer_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := 1
defer func() { __check(fmt.Sprint(value), "3") }()
value = 3
}
