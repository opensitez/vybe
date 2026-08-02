// vybe-test: go/defer_panic_recover_extra/defer_with_multiple_returns_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func build() int { defer func() {}()
if true { return 1 }
return 2 }
func main() { _ = build }
