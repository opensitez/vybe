// vybe-test: kotlin/tuples/test_tuple_triple_equality
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Triple(1, 2, 3) == Triple(1, 2, 3)).toString(), "true")
            __check((Triple(1, 2, 3) == Triple(3, 2, 1)).toString(), "false")
        }
