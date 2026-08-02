// vybe-test: go/composite_literal_keys/array_keyed_inside_keyed_struct_field
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type box struct { data [4]int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b := box{data: [4]int{0: 7, 3: 9}}
__check(fmt.Sprint(b.data[0]), "7")
__check(fmt.Sprint(b.data[2]), "0")
__check(fmt.Sprint(b.data[3]), "9")
}
