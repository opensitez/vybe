// vybe-test: kotlin/function_overloads/test_overload_with_nullable_non_nullable
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun show(v: String): String = "S:" + v
        fun show(v: String?): String = "N:" + (v ?: "nil")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((show("x")).toString(), "S:x")
        }
