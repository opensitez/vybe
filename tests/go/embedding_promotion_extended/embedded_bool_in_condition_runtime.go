// vybe-test: go/embedding_promotion_extended/embedded_bool_in_condition_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct { ok bool }
type outer struct { inner }
func main() { o := outer{inner: inner{ok: true}}
if o.ok { fmt.Println(1) } else { fmt.Println(0) } }
