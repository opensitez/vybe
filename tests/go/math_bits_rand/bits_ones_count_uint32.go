// vybe-test: go/math_bits_rand/bits_ones_count_uint32
// origin: languages/go/tests/go/test_math_bits_rand.rs
// vybe-test-mode: compile

package main
import "math/bits"
func main() { _ = bits.OnesCount32(0xFFFFFFFF) }
