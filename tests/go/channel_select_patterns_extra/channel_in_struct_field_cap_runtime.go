// vybe-test: go/channel_select_patterns_extra/channel_in_struct_field_cap_runtime
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs

package main
import "fmt"
type holder struct { ch chan int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := holder{ch: make(chan int, 5)}
__check(fmt.Sprint(cap(value.ch)), "5")
}
