// vybe-test: kotlin/conversions/test_string_to_int_radix_variants
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("1010".toInt(2)).toString(), "10")
            __check(("ff".toInt(16)).toString(), "255")
            __check(("-11".toInt(3)).toString(), "-4")
            __check(("77".toInt(8)).toString(), "63")
        }
