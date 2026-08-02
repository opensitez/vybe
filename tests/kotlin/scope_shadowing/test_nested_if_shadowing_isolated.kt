// vybe-test: kotlin/scope_shadowing/test_nested_if_shadowing_isolated
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n = 10
            fun test(x: Int): Int {
                val n = x + 1
                return if (x > 5) {
                    val n = n + 5
                    n
                } else {
                    n
                }
            }
            __check((test(6)).toString(), "12")
            __check((n).toString(), "10")
        }
