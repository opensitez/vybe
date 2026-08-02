// vybe-test: kotlin/tuples/test_tuple_triple_creation_and_components
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val triple = Triple(1, "x", true)
            __check((triple.first).toString(), "1")
            __check((triple.second).toString(), "x")
            __check((triple.third).toString(), "true")
        }
