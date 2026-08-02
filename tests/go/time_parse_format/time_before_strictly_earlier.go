// vybe-test: go/time_parse_format/time_before_strictly_earlier
// origin: languages/go/tests/go/test_time_parse_format.rs

package main
import "fmt"
import "time"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { early := time.Unix(10, 0)
late := time.Unix(20, 0)
__check(fmt.Sprint(early.Before(late)), "true") }
