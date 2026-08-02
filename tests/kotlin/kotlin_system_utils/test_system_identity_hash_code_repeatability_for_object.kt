// vybe-test: kotlin/kotlin_system_utils/test_system_identity_hash_code_repeatability_for_object
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = arrayOf(1, 2, 3)
            val first = System.identityHashCode(value)
            val second = System.identityHashCode(value)
            __check((first != 0).toString(), "true")
            __check((second != 0).toString(), "true")
            __check((first == second).toString(), "true")
        }
