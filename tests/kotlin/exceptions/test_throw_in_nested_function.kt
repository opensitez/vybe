// vybe-test: kotlin/exceptions/test_throw_in_nested_function
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun fail() {
            throw Exception("inner")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                fail()
            } catch (e: Exception) {
                __check(("caught").toString(), "caught")
            }
        }
