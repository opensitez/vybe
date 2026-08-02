// vybe-test: kotlin/scope/test_local_function_captures_enclosing_scope
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 0
            fun add(step: Int) {
                total += step
            }
            add(3)
            add(4)
            __check((total).toString(), "7")
        }
