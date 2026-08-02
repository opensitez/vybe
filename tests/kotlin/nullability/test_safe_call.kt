// vybe-test: kotlin/nullability/test_safe_call
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

class User(val name: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val u: User? = null
            __check((u?.name ?: "No User").toString(), "No User")
        }
