// vybe-test: kotlin/tuples/test_tuple_triple_to_string
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val triple = Triple(1, 2, 3)
            __check((triple).toString(), "(1, 2, 3)")
        }
