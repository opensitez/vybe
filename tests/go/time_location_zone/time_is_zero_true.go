// vybe-test: go/time_location_zone/time_is_zero_true
// origin: languages/go/tests/go/test_time_location_zone.rs

package main
import "fmt"
import "time"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var t time.Time
__check(fmt.Sprint(t.IsZero()), "true") }
