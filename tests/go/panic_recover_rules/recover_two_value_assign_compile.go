// vybe-test: go/panic_recover_rules/recover_two_value_assign_compile
// origin: languages/go/tests/go/test_panic_recover_rules.rs
// vybe-test-mode: compile

package main
func run() { defer func() { v, ok := recover().(int)
_ = v
_ = ok }()
panic(1) }
func main() { run() }
