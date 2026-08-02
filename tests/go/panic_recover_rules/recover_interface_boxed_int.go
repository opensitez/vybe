// vybe-test: go/panic_recover_rules/recover_interface_boxed_int
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { v := recover().(interface{})
fmt.Println(v.(int)) }()
panic(interface{}(5)) }
func main() { run() }
