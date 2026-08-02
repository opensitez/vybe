// vybe-test: kotlin/numeric_types/test_increment_prefix_and_postfix
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var count = 5
            __check((++count).toString(), "6")
            __check((count).toString(), "6")
            __check((count++).toString(), "6")
            __check((count).toString(), "7")
        }
