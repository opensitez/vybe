// vybe-test: kotlin/try_catch_flow/test_try_with_boolean_guard
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun guard(x: Int): Boolean = x > 0
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = try {
                if (guard(1)) 10 else throw Exception("bad")
            } catch (e: Exception) {
                0
            }
            __check((x).toString(), "10")
        }
