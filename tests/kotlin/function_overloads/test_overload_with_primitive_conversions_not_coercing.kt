// vybe-test: kotlin/function_overloads/test_overload_with_primitive_conversions_not_coercing
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun decode(v: Int): String = "i"
        fun decode(v: String): String = "s"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((decode(1)).toString(), "i")
            __check((decode("1")).toString(), "s")
        }
