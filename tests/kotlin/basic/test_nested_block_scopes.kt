// vybe-test: kotlin/basic/test_nested_block_scopes
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 1
            {
                val y = 2
                __check((x + y).toString(), "3")
            }
            __check((x).toString(), "1")
        }
