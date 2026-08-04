// vybe-test: go/time_location_zone/time_date_year_boundary
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

func main() { t := time.Date(1999, time.December, 31, 23, 59, 59, 0, time.UTC)
__p(fmt.Sprint(t.Year()))
__p(fmt.Sprint(t.Month())) 
__check("1999\nDecember")
}
