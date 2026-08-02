// vybe-test: kotlin/type_casts/test_is_type_check
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val str = "hello"
            if (str is String) {
                __check(("is string").toString(), "is string")
            }
            if (str !is Int) {
                __check(("not int").toString(), "not int")
            }
        }
