// vybe-test: kotlin/receiver_this_context/test_this_with_label_in_nested_functions
// origin: languages/kotlin/tests/kotlin/test_receiver_this_context.rs

class Container {
            val name = "container"

            fun make(prefix: String): String {
                fun nested() = this@Container.name + prefix
                return nested()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Container().make("X")).toString(), "containerX")
        }
