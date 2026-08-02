// vybe-test: kotlin/try_catch_flow/test_try_catch_with_custom_error
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

class AppError(msg: String) : Exception(msg)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                throw AppError("boom")
            } catch (e: AppError) {
                __check((e.message).toString(), "boom")
            }
        }
