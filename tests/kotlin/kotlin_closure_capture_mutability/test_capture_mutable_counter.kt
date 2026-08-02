// vybe-test: kotlin/kotlin_closure_capture_mutability/test_capture_mutable_counter
// origin: languages/kotlin/tests/kotlin/test_kotlin_closure_capture_mutability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var count = 0
            val inc = { count += 1 }
            inc()
            inc()
            __check((count).toString(), "2")
        }
