// vybe-test: kotlin/conversions/test_string_to_long_radix_negative_and_large
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("7fffffff".toLong(16)).toString(), "2147483647")
            __check(("-100000000".toLong(2)).toString(), "-256")
            __check(("1fffffffffffff".toLong(16)).toString(), "9007199254740991")
        }
