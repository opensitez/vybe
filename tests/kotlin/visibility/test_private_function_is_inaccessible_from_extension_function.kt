// vybe-test: kotlin/visibility/test_private_function_is_inaccessible_from_extension_function
// origin: languages/kotlin/tests/kotlin/test_visibility.rs
// vybe-test-mode: compile

class Item {
            private fun secret(): String = "x"
        }

        fun Item.exposed(): String = secret()

        fun main() {
            println(Item().exposed())
        }

