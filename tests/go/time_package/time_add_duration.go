// vybe-test: go/time_package/time_add_duration
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

func main() { t := time.Unix(100, 0)
later := t.Add(10 * time.Second)
__check(fmt.Sprint(later.Unix()), "110") }
