// vybe-test: kotlin/scope_shadowing/test_shadowing_in_destructuring
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val outer = "outer"
            val pair = Pair("inner", 1)
            val (value, count) = pair
            __check((value).toString(), "inner")
            __check((count).toString(), "1")
            __check((outer).toString(), "outer")
        }
