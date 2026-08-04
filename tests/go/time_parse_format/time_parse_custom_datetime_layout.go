// vybe-test: go/time_parse_format/time_parse_custom_datetime_layout
// origin: languages/go/tests/go/test_time_parse_format.rs

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

func main() { t, _ := time.Parse("2006-01-02 15:04:05", "2020-01-02 03:04:05")
__p(fmt.Sprint(t.Hour())) 
__check("3")
}
