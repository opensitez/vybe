// vybe-test: kotlin/kotlin_type_cast_edge/test_forced_cast_to_expected_type
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_cast_edge.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val any: Any = 42
            val value = any as Int
            __check((value + 1).toString(), "43")
        }
