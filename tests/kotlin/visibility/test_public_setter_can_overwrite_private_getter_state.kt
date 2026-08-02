// vybe-test: kotlin/visibility/test_public_setter_can_overwrite_private_getter_state
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Item {
            private var value = "x"
            var display: String
                get() = value
                private set(next) {
                    value = next
                }

            fun reset(next: String) {
                display = next
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Item()
            __check((item.display).toString(), "x")
            item.reset("y")
            __check((item.display).toString(), "y")
        }
