// vybe-test: kotlin/visibility/test_protected_property_not_visible_to_unrelated_type
// origin: languages/kotlin/tests/kotlin/test_visibility.rs
// vybe-test-mode: compile

open class Base {
            protected var value: Int = 1
        }

        fun leak(base: Base) {
            println(base.value)
        }

        fun main() {
            leak(Base())
        }

