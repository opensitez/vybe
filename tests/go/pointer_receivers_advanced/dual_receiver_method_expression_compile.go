// vybe-test: go/pointer_receivers_advanced/dual_receiver_method_expression_compile
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs
// vybe-test-mode: compile

package main
type ledger struct { balance int }
func (l ledger) snapshot() int { return l.balance }
func (l *ledger) deposit(v int) { l.balance += v }
func main() { _ = ledger.snapshot
_ = (*ledger).deposit }
