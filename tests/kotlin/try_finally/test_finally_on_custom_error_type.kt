// vybe-test: kotlin/try_finally/test_finally_on_custom_error_type
// origin: languages/kotlin/tests/kotlin/test_try_finally.rs

class Err : Exception()
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    try {
        throw Err()
    } catch (e: Err) {
        __check(("err").toString(), "err")
    } finally {
        __check(("fin").toString(), "fin")
    }
}
