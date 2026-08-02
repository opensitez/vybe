// vybe-test: go/panic_recover_rules/panic_struct_field_via_recover
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
type err struct { code int }
func run() { defer func() { e := recover().(err)
fmt.Println(e.code) }()
panic(err{code: 5}) }
func main() { run() }
