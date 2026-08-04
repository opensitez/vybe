// vybe-test: go/composite_literals_extra/array_literal_with_indexed_elements_runtime
// origin: languages/go/tests/go/test_composite_literals_extra.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { values := [4]int{1: 7, 3: 9}
fmt.Println(values[1])
fmt.Println(values[3])
}
