// vybe-test: kotlin/visibility/test_private_constructor_prevents_external_call
// origin: languages/kotlin/tests/kotlin/test_visibility.rs
// vybe-test-mode: compile

class Guard private constructor(val value: Int)

        fun main() {
            val value = Guard(1)
            println(value.value)
        }

