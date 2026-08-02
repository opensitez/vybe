// vybe-test: go/defer_panic_variants/defer_inner_frame_finishes_before_outer
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func inner() { defer __check(fmt.Sprint("inner"), "inner")
}
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer __check(fmt.Sprint("outer"), "outer")
inner()
}
