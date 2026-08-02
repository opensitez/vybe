// vybe-test: kotlin/operators/test_in_operator_on_empty_range_and_iterable
// origin: languages/kotlin/tests/kotlin/test_operators.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val empty = 1..0
            __check((empty.isEmpty()).toString(), "true")
            __check((1 in empty).toString(), "false")
            val present = 5 in 1..10
            val absent = 11 in 1..10
            __check((present).toString(), "true")
            __check((absent).toString(), "false")
        }
