// vybe-test: kotlin/scope/test_local_function_reads_enclosing_immutables
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = 10
            fun bump(): Int {
                return base + 1
            }
            fun bumpTwice(): Int {
                return bump() + 1
            }
            __check((bump()).toString(), "11")
            __check((bumpTwice()).toString(), "12")
        }
