// vybe-test: go/iota_enumerations/iota_typed_constants
// origin: languages/go/tests/go/test_iota_enumerations.rs

package main
import "fmt"
type status int
const ( Ok status = iota; Err; Unknown )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(int(Ok)), "0")
__check(fmt.Sprint(int(Unknown)), "2") }
