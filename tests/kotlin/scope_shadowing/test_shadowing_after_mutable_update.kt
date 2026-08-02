// vybe-test: kotlin/scope_shadowing/test_shadowing_after_mutable_update
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value = 1
            run {
                var value = value + 1
                __check((value).toString(), "2")
            }
            __check((value).toString(), "1")
        }
