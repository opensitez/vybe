// vybe-test: kotlin/basic/test_compound_mod_assign
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var x = 29
            x %= 6
            __check((x).toString(), "5")
        }
