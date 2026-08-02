// vybe-test: go/pointer_receivers_advanced/address_of_field_passed_to_mutator_compile
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs
// vybe-test-mode: compile

package main
func scale(target *int, factor int) { *target = *target * factor }
func main() { value := struct{ n int }{n: 3}
scale(&value.n, 4) }
