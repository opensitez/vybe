// vybe-test: go/os_process_environ/os_getenv_with_default_returns_set_value
// origin: languages/go/tests/go/test_os_process_environ.rs

package main
import "fmt"
import "os"
func getenvDefault(key, def string) string { if v := os.Getenv(key); v != "" { return v }
return def }
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

func main() { os.Setenv("VYBE_TMP_TEST_KEY", "live")
__p(fmt.Sprint(getenvDefault("VYBE_TMP_TEST_KEY", "fallback"))) 
__check("live")
}
