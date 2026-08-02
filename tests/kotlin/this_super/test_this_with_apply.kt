// vybe-test: kotlin/this_super/test_this_with_apply
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val out = StringBuilder().apply { this.append("a") }.toString()
        __check((out).toString(), "a")
    }
