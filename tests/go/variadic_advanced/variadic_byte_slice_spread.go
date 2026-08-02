// vybe-test: go/variadic_advanced/variadic_byte_slice_spread
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func total(bytes ...byte) int { t := 0
for _, b := range bytes { t += int(b) }
return t }
func main() { data := []byte{'a', 'b'}
fmt.Println(total(data...)) }
