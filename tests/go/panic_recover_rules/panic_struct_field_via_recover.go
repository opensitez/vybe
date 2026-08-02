// vybe-test: go/panic_recover_rules/panic_struct_field_via_recover
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
type err struct { code int }
func run() { defer func() { e := recover().(err)
__check(fmt.Sprint(e.code), "5") }()
panic(err{code: 5}) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
