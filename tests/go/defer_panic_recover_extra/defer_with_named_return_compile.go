// vybe-test: go/defer_panic_recover_extra/defer_with_named_return_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func build() (result int) { defer func() { result++ }()
return 1 }
func main() { _ = build }
