// vybe-test: go/channel_direction_extended/send_only_in_map_value
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
m := map[string]chan<- int{"k": ch}
m["k"] <- 17
__check(fmt.Sprint(<-ch), "17") }
