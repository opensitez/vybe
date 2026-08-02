// vybe-test: go/composite_literal_keys/anonymous_struct_literal_keyed_fields
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { v := struct { id int
name string }{name: "go", id: 9}
__check(fmt.Sprint(v.id), "9")
__check(fmt.Sprint(v.name), "go")
}
