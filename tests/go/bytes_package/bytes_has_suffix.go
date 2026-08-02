// vybe-test: go/bytes_package/bytes_has_suffix
// origin: languages/go/tests/go/test_bytes_package.rs

package main
import "fmt"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(bytes.HasSuffix([]byte("golang"), []byte("lang"))), "true") }
