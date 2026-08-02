// vybe-test: go/pointer_receivers_advanced/new_and_literal_pointer_receiver_compile
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs
// vybe-test-mode: compile

package main
type tally struct { sum int }
func (t *tally) add(v int) { t.sum += v }
func main() { a := new(tally)
b := &tally{}
a.add(1)
b.add(2) }
