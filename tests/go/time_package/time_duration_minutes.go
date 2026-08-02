// vybe-test: go/time_package/time_duration_minutes
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

func main() { d := 2 * time.Minute
__check(fmt.Sprint(d.Minutes()), "2") }
