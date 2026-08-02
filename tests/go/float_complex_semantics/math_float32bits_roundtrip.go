// vybe-test: go/float_complex_semantics/math_float32bits_roundtrip
// origin: languages/go/tests/go/test_float_complex_semantics.rs

package main
import "fmt"
import "math"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { bits := math.Float32bits(2.5)
__check(fmt.Sprint(math.Float32frombits(bits)), "2.5") }
