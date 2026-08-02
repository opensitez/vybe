// vybe-test: go/defer_panic_recover_extra/defer_modify_named_result_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func build() (result int) { defer func() { result += 2 }()
return 1 }
func main() { _ = build }
