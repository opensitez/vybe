// vybe-test: kotlin/function_overloads/test_overload_boolean_and_unit
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun marker(v: Int): Int = v
        fun marker(v: Boolean): String = if (v) "on" else "off"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((marker(2)).toString(), "2")
            __check((marker(false)).toString(), "off")
        }
