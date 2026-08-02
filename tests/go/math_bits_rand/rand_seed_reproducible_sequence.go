// vybe-test: go/math_bits_rand/rand_seed_reproducible_sequence
// origin: languages/go/tests/go/test_math_bits_rand.rs
// vybe-test-mode: compile

package main
import "fmt"
import "math/rand"
func main() { rand.Seed(1)
fmt.Println(rand.Intn(10))
rand.Seed(1)
fmt.Println(rand.Intn(10)) }
