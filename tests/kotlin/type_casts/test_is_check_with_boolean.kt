// vybe-test: kotlin/type_casts/test_is_check_with_boolean
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val flag = true
            if (flag is Boolean) {
                __check(("is boolean").toString(), "is boolean")
            }
        }
