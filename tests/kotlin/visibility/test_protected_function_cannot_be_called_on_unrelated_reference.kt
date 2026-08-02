// vybe-test: kotlin/visibility/test_protected_function_cannot_be_called_on_unrelated_reference
// origin: languages/kotlin/tests/kotlin/test_visibility.rs
// vybe-test-mode: compile

open class Base {
            protected fun label(): String = "base"
        }

        fun call(base: Base): String = base.label()

        fun main() {
            println(call(Base()))
        }

