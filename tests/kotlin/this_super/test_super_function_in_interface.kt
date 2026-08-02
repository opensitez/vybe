// vybe-test: kotlin/this_super/test_super_function_in_interface
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

interface I { fun tag() = "i" }
open class A { open fun tag() = "a" }
class B: A(), I { override fun tag() = super<A>.tag() + super<I>.tag() }

