// vybe-test: go/math_bits_rand/rand_seed_before_multiple_intn
// origin: languages/go/tests/go/test_math_bits_rand.rs
// vybe-test-mode: compile

package main
import "math/rand"
func main() { rand.Seed(99)
_ = rand.Intn(10)
_ = rand.Intn(10) }
