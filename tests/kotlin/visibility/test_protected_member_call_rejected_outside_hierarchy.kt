// vybe-test: kotlin/visibility/test_protected_member_call_rejected_outside_hierarchy
// origin: languages/kotlin/tests/kotlin/test_visibility.rs
// vybe-test-mode: compile

open class Base {
            protected fun label(): String = "base"
        }

        class NotChild : Base()

        fun main() {
            val value = Base()
            println(value.label())
        }

