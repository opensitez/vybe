// vybe-test: go/structs/struct_ambiguous_promotion_at_equal_depth_is_rejected
// origin: languages/go/tests/go/test_structs.rs
// vybe-test-mode: compile

package main
import "fmt"
type A struct { ID int }
type B struct { ID int }
type C struct { A
B }
func main() { c := C{}
fmt.Println(c.ID) }
