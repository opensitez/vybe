// vybe-test: go/type_conversions_extra/struct_field_conversion_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
type holder struct { count int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := holder{count: 12}
__check(fmt.Sprint(float64(value.count)), "12")
}
