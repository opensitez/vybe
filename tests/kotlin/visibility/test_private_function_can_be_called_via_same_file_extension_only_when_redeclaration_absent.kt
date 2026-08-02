// vybe-test: kotlin/visibility/test_private_function_can_be_called_via_same_file_extension_only_when_redeclaration_absent
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Item {
            private fun secret(): String = "inside"
        }

        fun itemSecret(item: Item): String = item.access()

        private fun Item.access(): String = secret()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((itemSecret(Item())).toString(), "inside")
        }
