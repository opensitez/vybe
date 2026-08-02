// vybe-test: go/constants_iota_advanced/iota_storage_kb_mb_gb
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( KB = 1 << (10 * iota); MB; GB )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(KB), "1")
__check(fmt.Sprint(MB), "1048576") }
