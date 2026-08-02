// vybe-test: kotlin/this_super/test_this_in_run_return
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x = StringBuilder().run {
            this.append("x")
            this.toString()
        }
        __check((x).toString(), "x")
    }
