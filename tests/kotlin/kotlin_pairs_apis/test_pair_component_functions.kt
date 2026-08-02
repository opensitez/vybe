// vybe-test: kotlin/kotlin_pairs_apis/test_pair_component_functions
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 7 to 9
            __check((value.component1()).toString(), "7")
            __check((value.component2()).toString(), "9")
        }
