// vybe-test: go/generics_constraints_extended/generic_comparable_chan_key
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
func KeyType[K comparable]() K { var z K
return z }
func main() { _ = KeyType[chan int]() }
