// vybe-test: go/functions_patterns_extra/named_return_explicit_assignment_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func classify(v int) (label string) { if v > 0 { label = "pos"
return }
label = "zero"
return }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(classify(3)), "pos")
}
