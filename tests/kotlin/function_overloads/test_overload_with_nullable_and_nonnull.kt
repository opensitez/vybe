// vybe-test: kotlin/function_overloads/test_overload_with_nullable_and_nonnull
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun show(v: String): String = "NN"
        fun show(v: String?): String = "NULL"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((show("x")).toString(), "NN")
            __check((show(null)).toString(), "NULL")
        }
