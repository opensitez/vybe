// vybe-test: kotlin/function_types/test_function_type_noarg_constructor
// origin: languages/kotlin/tests/kotlin/test_function_types.rs

val defaultFactory: () -> String = { "ok" }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((defaultFactory()).toString(), "ok")
        }
