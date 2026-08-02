// vybe-test: kotlin/type_casts/test_multiple_type_checks
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun check(value: Any) {
            if (value is String) {
                __check(("string").toString(), "string")
            } else if (value !is Int) {
                __check(("not int").toString(), "int")
            } else {
                __check(("int").toString(), "not int")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            check("x")
            check(3)
            check(true)
        }
