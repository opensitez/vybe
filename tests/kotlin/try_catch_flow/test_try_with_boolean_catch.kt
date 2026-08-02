// vybe-test: kotlin/try_catch_flow/test_try_with_boolean_catch
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ok = try {
                true
            } catch (e: Exception) {
                false
            }
            __check((ok).toString(), "true")
        }
