// vybe-test: kotlin/basic/test_boolean_complex
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = true
            val b = false
            val c = true
            __check(((a || b) && c).toString(), "true")
        }
