// vybe-test: go/label_break_continue/labeled_continue_on_switch_inside_loop
// origin: languages/go/tests/go/test_label_break_continue.rs
// vybe-test-mode: compile

package main
func main() { loop: for i := 0; i < 2; i++ { switch i { case 0: continue loop } } }
