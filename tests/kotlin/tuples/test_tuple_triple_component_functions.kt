// vybe-test: kotlin/tuples/test_tuple_triple_component_functions
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val triple = Triple("a", 2, 3)
            __check((triple.component1()).toString(), "a")
            __check((triple.component2()).toString(), "2")
            __check((triple.component3()).toString(), "3")
        }
