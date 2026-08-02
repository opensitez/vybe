// vybe-test: go/panic_recover_rules/re_panic_after_recover_compile
// origin: languages/go/tests/go/test_panic_recover_rules.rs
// vybe-test-mode: compile

package main
func run() { defer func() { recover()
panic("again") }()
panic("first") }
func main() { run() }
