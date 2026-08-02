// vybe-test: kotlin/object_declarations/test_object_with_init_like_function_call_order_is_deterministic
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Log {
            val value: Int
            init {
                value = 10
            }

            fun value(): Int = value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Log.value).toString(), "10")
            __check((Log.value()).toString(), "10")
        }
