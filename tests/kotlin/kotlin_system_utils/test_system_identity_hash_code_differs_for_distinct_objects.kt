// vybe-test: kotlin/kotlin_system_utils/test_system_identity_hash_code_differs_for_distinct_objects
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Any()
            val b = Any()
            val first = System.identityHashCode(a)
            val second = System.identityHashCode(b)
            __check((first == second).toString(), "false")
            __check((first != 0).toString(), "true")
            __check((second != 0).toString(), "true")
        }
