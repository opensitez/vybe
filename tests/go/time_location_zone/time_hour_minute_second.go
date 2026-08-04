// vybe-test: go/time_location_zone/time_hour_minute_second
// origin: languages/go/tests/go/test_time_location_zone.rs

package main
import "fmt"
import "time"
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { t := time.Date(2020, 6, 15, 9, 8, 7, 0, time.UTC)
__p(fmt.Sprint(t.Hour()))
__p(fmt.Sprint(t.Minute()))
__p(fmt.Sprint(t.Second())) 
__check("9\n8\n7")
}
