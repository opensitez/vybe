// vybe-test: kotlin/scope/test_scope_limited_block_variable_lifetime
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value = 1
            if (value == 1) {
                val block = 8
                value += block
            }
            __check((value).toString(), "9")
        }
