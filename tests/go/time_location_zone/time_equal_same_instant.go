// vybe-test: go/time_location_zone/time_equal_same_instant
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

func main() { a := time.Unix(100, 0)
b := time.Unix(100, 0)
__check(fmt.Sprint(a.Equal(b)), "true") }
