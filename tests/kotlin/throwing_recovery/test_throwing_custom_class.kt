// vybe-test: kotlin/throwing_recovery/test_throwing_custom_class
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

class DomainError(message: String) : Exception(message)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                throw DomainError("oops")
            } catch (e: DomainError) {
                __check((e.message).toString(), "oops")
            }
        }
