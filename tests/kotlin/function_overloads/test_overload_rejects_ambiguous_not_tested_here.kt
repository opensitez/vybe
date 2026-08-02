// vybe-test: kotlin/function_overloads/test_overload_rejects_ambiguous_not_tested_here
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun tag(v: Int, suffix: String = "a"): String = "i" + suffix
        fun tag(v: Double, suffix: String = "b"): String = "d" + suffix
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((tag(1)).toString(), "ia")
            __check((tag(1.0, "Z")).toString(), "dZ")
        }
