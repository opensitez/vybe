// vybe-test: kotlin/visibility/test_private_constructor_is_restricted_to_same_scope
// origin: languages/kotlin/tests/kotlin/test_visibility.rs
// vybe-test-mode: compile

class Core private constructor(val value: String) {
            companion object {
                fun from(value: String): Core = Core(value)
            }
        }

        class Factory {
            fun create(value: String): Core {
                return Core(value)
            }
        }

        fun main() {
            println(Factory().create("x"))
        }

