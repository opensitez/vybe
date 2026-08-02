// vybe-test: go/defer_panic_variants/recover_captures_struct_field_via_type_assertion
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
type stop struct { code int }
func run() { defer func() { value := recover()
if err, ok := value.(stop); ok { fmt.Println(err.code) } }()
panic(stop{code: 42}) }
func main() { run() }
