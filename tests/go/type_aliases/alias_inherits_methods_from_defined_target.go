// vybe-test: go/type_aliases/alias_inherits_methods_from_defined_target
// origin: languages/go/tests/go/test_type_aliases.rs
// vybe-test-mode: compile

package main
type Units int
func (u Units) sign() int { if u < 0 { return -1 }
if u > 0 { return 1 }
return 0 }
type Reading = Units
func main() { _ = Reading(3).sign() }
