// vybe-test: go/math_bits_rand/rand_intn_loop_accumulates_valid
// origin: languages/go/tests/go/test_math_bits_rand.rs

package main
import "fmt"
import "math/rand"
func main() { ok := 0
i := 0
for i < 4 { if rand.Intn(3) < 3 { ok++ }
i++ }
fmt.Println(ok) }
