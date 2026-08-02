// vybe-test: go/time_package/time_before_after
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

func main() { early := time.Unix(1,0)
late := time.Unix(2,0)
__check(fmt.Sprint(early.Before(late)), "true")
__check(fmt.Sprint(late.After(early)), "true") }
