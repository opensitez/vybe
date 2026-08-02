// vybe-test: kotlin/numeric_types/test_numeric_relational_operator_network
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((3 > 2).toString(), "true")
            __check((3 >= 3).toString(), "true")
            __check((3 < 4.0).toString(), "true")
            __check((3L != 4L).toString(), "true")
            __check((3.0 == 3).toString(), "true")
        }
