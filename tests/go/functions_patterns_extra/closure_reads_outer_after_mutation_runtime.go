// vybe-test: go/functions_patterns_extra/closure_reads_outer_after_mutation_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { prefix := "go"
fn := func() string { return prefix }
prefix = "vybe"
__check(fmt.Sprint(fn()), "vybe")
}
