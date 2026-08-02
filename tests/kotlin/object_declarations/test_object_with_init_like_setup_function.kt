// vybe-test: kotlin/object_declarations/test_object_with_init_like_setup_function
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Registry {
            var value = 0
            fun setup() {
                value = 3
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Registry.setup()
            __check((Registry.value).toString(), "3")
        }
