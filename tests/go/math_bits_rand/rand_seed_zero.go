// vybe-test: go/math_bits_rand/rand_seed_zero
// origin: languages/go/tests/go/test_math_bits_rand.rs
// vybe-test-mode: compile

package main
import "math/rand"
func main() { rand.Seed(0)
_ = rand.Intn(5) }
