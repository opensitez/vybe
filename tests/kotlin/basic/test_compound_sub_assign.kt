// vybe-test: kotlin/basic/test_compound_sub_assign
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var x = 20
            x -= 8
            __check((x).toString(), "12")
        }
