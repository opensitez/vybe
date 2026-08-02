// vybe-test: kotlin/throwing_recovery/test_throwing_in_nested_call
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

class Boom : Exception("boom")
        fun explode() = throw Boom()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                explode()
            } catch (e: Boom) {
                __check((e.message).toString(), "boom")
            }
        }
