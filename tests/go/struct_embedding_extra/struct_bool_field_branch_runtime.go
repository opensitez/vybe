// vybe-test: go/struct_embedding_extra/struct_bool_field_branch_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type state struct { ok bool }
func main() { value := state{ok: true}
if value.ok { fmt.Println(1) } else { fmt.Println(0) } }
