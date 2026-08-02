// vybe-test: go/time_package/time_sub_duration
// origin: languages/go/tests/go/test_time_package.rs

package main
import "fmt"
import "time"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := time.Unix(100, 0)
b := time.Unix(40, 0)
__check(fmt.Sprint(a.Sub(b).Seconds()), "60") }
