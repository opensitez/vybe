// vybe-test: go/blank_identifier_extended/blank_for_init_discard
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func main() { total := 0
for _ = range 3 { total++ }
fmt.Println(total) }
