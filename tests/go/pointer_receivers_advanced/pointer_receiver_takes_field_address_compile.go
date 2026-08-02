// vybe-test: go/pointer_receivers_advanced/pointer_receiver_takes_field_address_compile
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs
// vybe-test-mode: compile

package main
type holder struct { gauge int }
func (h *holder) reset() { target := &h.gauge
*target = 0 }
func main() { value := &holder{gauge: 3}
value.reset() }
