// vybe-test: go/struct_embedding_advanced/embedded_bool_promoted_in_condition_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type inner struct { ready bool }
type outer struct { inner }
func main() { value := outer{inner: inner{ready: true}}
if value.ready { fmt.Println(1) } else { fmt.Println(0) } }
