// vybe-test: kotlin/type_casts/test_cast_with_function_return
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun toText(value: Any): String {
            val casted = value as? String
            return casted ?: "fallback"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((toText("hello")).toString(), "hello")
            __check((toText(2)).toString(), "fallback")
        }
