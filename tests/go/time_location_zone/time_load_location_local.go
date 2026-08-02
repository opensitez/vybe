// vybe-test: go/time_location_zone/time_load_location_local
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

func main() { loc, err := time.LoadLocation("Local")
__check(fmt.Sprint(err == nil), "true")
__check(fmt.Sprint(loc != nil), "true") }
