// vybe-test: kotlin/operators/test_pre_increment_and_post_increment_difference
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var a = 5
            __check((++a).toString(), "6")
            __check((a).toString(), "6")
            __check((a++).toString(), "6")
            __check((a).toString(), "7")
        }
