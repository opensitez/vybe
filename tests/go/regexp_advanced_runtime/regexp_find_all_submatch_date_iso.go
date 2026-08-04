// vybe-test: go/regexp_advanced_runtime/regexp_find_all_submatch_date_iso
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs

package main
import "fmt"
import "regexp"
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

func main() { re := regexp.MustCompile(`(\d{4})-(\d{2})-(\d{2})`)
m := re.FindAllStringSubmatch("2024-06-30", -1)
__p(fmt.Sprint(m[0][1]))
__p(fmt.Sprint(m[0][3])) 
__check("2024\n30")
}
