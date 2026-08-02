// vybe-test: kotlin/basic/test_block_scoped_shadowing_and_assignment
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 10
            run {
                val x = 99
                __check((x).toString(), "99")
            }
            __check((x).toString(), "10")
        }
