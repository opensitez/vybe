// vybe-test: go/math_big_int/big_float_parse
// origin: languages/go/tests/go/test_math_big_int.rs
// vybe-test-mode: compile

package main
import "math/big"
func main() { f := new(big.Float)
_, _ = f.SetString("1.23e2") }
