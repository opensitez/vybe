// vybe-test: kotlin/kotlin_system_utils/test_measure_time_and_runtime_identity_in_one_block
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val elapsed = kotlin.system.measureTimeMillis {
                val id1 = System.identityHashCode(Any())
                val id2 = System.identityHashCode(Any())
                __check((id1 == id2).toString(), "false")
            }
            __check((elapsed >= 0).toString(), "true")
        }
