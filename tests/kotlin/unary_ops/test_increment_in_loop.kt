// vybe-test: kotlin/unary_ops/test_increment_in_loop
// origin: languages/kotlin/tests/kotlin/test_unary_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var x = 0
            var out = 0
            repeat(3) {
                out += ++x
            }
            __check((out).toString(), "6")
        }
