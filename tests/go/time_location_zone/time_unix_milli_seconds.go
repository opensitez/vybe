// vybe-test: go/time_location_zone/time_unix_milli_seconds
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

func main() { t := time.UnixMilli(1500)
__check(fmt.Sprint(t.Unix()), "1")
__check(fmt.Sprint(t.UnixMilli()), "1500") }
