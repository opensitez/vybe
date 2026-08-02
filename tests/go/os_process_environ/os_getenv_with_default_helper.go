// vybe-test: go/os_process_environ/os_getenv_with_default_helper
// origin: languages/go/tests/go/test_os_process_environ.rs

package main
import "fmt"
import "os"
func getenvDefault(key, def string) string { if v := os.Getenv(key); v != "" { return v }
return def }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(getenvDefault("VYBE_MISSING_KEY_ABC", "fallback")), "fallback") }
