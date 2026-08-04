// vybe-test: go/regexp_advanced_runtime/regexp_subexp_names_named_group
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

func main() { re := regexp.MustCompile(`(?P<year>\d{4})`)
names := re.SubexpNames()
__p(fmt.Sprint(names[1]))
__p(fmt.Sprint(len(names))) 
__check("year\n2")
}
