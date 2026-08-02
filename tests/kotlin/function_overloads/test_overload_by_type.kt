// vybe-test: kotlin/function_overloads/test_overload_by_type
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun cast(v: Int): String = "I"
        fun cast(v: String): String = "S"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((cast(3)).toString(), "I")
            __check((cast("x")).toString(), "S")
        }
