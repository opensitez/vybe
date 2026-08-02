// vybe-test: go/channel_select_patterns_extra/channel_return_from_function_runtime
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs

package main
import "fmt"
func build() chan int { return make(chan int, 4) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(cap(build())), "4")
}
