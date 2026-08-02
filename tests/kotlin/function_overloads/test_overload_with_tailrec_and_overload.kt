// vybe-test: kotlin/function_overloads/test_overload_with_tailrec_and_overload
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun build(v: Int): Int = v
        fun build(v: Int, s: String): String = s + v
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((build(1)).toString(), "1")
            __check((build(1, "#")).toString(), "#1")
        }
