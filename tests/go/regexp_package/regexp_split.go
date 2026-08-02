// vybe-test: go/regexp_package/regexp_split
// origin: languages/go/tests/go/test_regexp_package.rs

package main
import "fmt"
import "regexp"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { re := regexp.MustCompile(`[,\s]+`)
parts := re.Split("a, b  c", -1)
__check(fmt.Sprint(len(parts)), "3")
__check(fmt.Sprint(parts[2]), "c") }
