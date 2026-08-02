// vybe-test: go/type_conversions_extra/conversion_in_if_compare_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
func main() { if int(21.1) == 21 { fmt.Println(1) } else { fmt.Println(0) } }
