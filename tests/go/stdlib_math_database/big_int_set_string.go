// vybe-test: go/stdlib_math_database/big_int_set_string
// origin: languages/go/tests/go/test_stdlib_math_database.rs
// vybe-test-mode: compile

package main
import "math/big"
func main() { z := new(big.Int)
_, _ = z.SetString("ff", 16) }
