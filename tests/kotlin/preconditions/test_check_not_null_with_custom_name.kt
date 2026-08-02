// vybe-test: kotlin/preconditions/test_check_not_null_with_custom_name
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun read(v: String?): String {
                return checkNotNull(v)
            }
            __check((read("x")).toString(), "x")
        }
