// vybe-test: go/panic_recover_rules/recover_assigned_to_wrong_type_compile
// origin: languages/go/tests/go/test_panic_recover_rules.rs
// vybe-test-mode: compile

package main
func run() { defer func() { s := recover().(string)
_ = s }()
panic(1) }
func main() { run() }
