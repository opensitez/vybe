// vybe-test: kotlin/boolean_logic/test_boolean_ordering_via_to_int_like_comparison
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = true
            val b = false
            __check((if (a == b) 0 else if (a && !b) 1 else -1).toString(), "1")
            __check((if (!a == b) "swap" else "noswap").toString(), "noswap")
        }
