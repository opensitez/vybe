// vybe-test: kotlin/basic/test_var_reassignment
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var x = 10
            x = 20
            __check((x).toString(), "20")
        }
