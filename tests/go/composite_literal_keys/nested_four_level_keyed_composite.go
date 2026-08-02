// vybe-test: go/composite_literal_keys/nested_four_level_keyed_composite
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type leaf struct { v int }
type branch struct { leaves []leaf }
type tree struct { parts []branch }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { tr := tree{parts: []branch{{leaves: []leaf{{v: 99}}}}}
__check(fmt.Sprint(tr.parts[0].leaves[0].v), "99")
}
