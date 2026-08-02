// vybe-test: go/control_flow_patterns_extra/goto_named_loop_exit_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { i := 0
Loop: if i == 1 { goto Exit }
i++
goto Loop
Exit: }
